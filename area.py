
# -*- coding: utf-8 -*-
# 地域貢献データ 年度×診療科 検索アプリ 再生成スクリプト
# 実行: python build_kouken_app.py
# 依存: pandas (openpyxl)
import pandas as pd
import json
from datetime import datetime
import os, re

EXCEL_FILE = "地域貢献_統合.xlsx"
HTML_FILE = "index.html"

# --- Excel読込（先頭シート） ---
xl = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
sheet_name = xl.sheet_names[0]
df = xl.parse(sheet_name)
df.columns = [str(c).strip() for c in df.columns]
for col in df.columns:
    df[col] = df[col].apply(lambda x: "" if pd.isna(x) else str(x))

# 任意：日付で昇順ソート（日付列があれば）
if "日付" in df.columns:
    def to_date(s: str):
        try:
            s2 = re.sub(r"[.年]", "/", s).replace("月", "/").replace("日", "")
            return pd.to_datetime(s2, errors="coerce")
        except Exception:
            return pd.NaT
    tmp_date = df["日付"].apply(to_date)
    df = df.assign(__sort_date=tmp_date).sort_values(by=["__sort_date", "日付"])
    df = df.drop(columns=["__sort_date"])

# --- 選択肢生成（独立：年度／診療科） ---
if ("年度" not in df.columns) or ("診療科" not in df.columns):
    raise ValueError("Excelに『年度』『診療科』列が必要です。")
years = sorted(list({y for y in df["年度"].tolist() if y}))
depts = sorted(list({d for d in df["診療科"].tolist() if d}))
choices = {"年度": years, "診療科": depts}

# --- 表示列の並び（主要項目を前に） ---
preferred = ["年度","事業所","診療科","発表者","日付","タイトル","主催/共催","形態","特記事項（年代、エリア限定等）"]
cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
records = df[cols].to_dict(orient="records")

# --- 見た目（CSS） ---
css = """
* { box-sizing: border-box; }
body { font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Hiragino Kaku Gothic Pro', 'Noto Sans JP', 'Yu Gothic', Meiryo, sans-serif; margin: 24px; }
h1 { font-size: 1.6rem; margin: 0 0 12px; }
header .meta { color: #666; font-size: .9rem; margin-bottom: 16px; }
.controls { display: flex; gap: 12px; flex-wrap: wrap; margin: 16px 0 12px; }
.controls label { font-weight: 600; font-size: .95rem; }
.card { border: 1px solid #ddd; border-radius: 8px; padding: 12px; margin: 12px 0; }
.card h2 { font-size: 1.2rem; margin: 0 0 8px; }
.count { color: #333; font-size: .95rem; margin-bottom: 8px; }
.tablewrap { overflow-x: auto; border: 1px solid #eee; border-radius: 6px; }
table { border-collapse: collapse; width: 100%; min-width: 960px; }
th, td { padding: 8px 10px; border-bottom: 1px solid #eee; text-align: left; white-space: nowrap; }
th { background: #f8f9fb; position: sticky; top: 0; z-index: 1; }
tr:nth-child(even) td { background: #fcfcff; }
.empty { color: #666; padding: 12px; }
.footer { margin-top: 18px; color: #555; font-size: .9rem; }
button { padding: 8px 12px; font-size: .9rem; border: 1px solid #ccc; border-radius: 6px; cursor: pointer; background: #fff; }
button:hover { background: #f4f5f7; }
.note { color: #777; font-size: .85rem; }
.badge { display:inline-block; padding: 2px 8px; background: #eef2ff; border:1px solid #c7d2fe; border-radius: 999px; font-size: .8rem; color:#1e40af; }

/* ▼ ここから dropdown（button+panel）＆縦並びチェックボックスに変更 */
.dropdown { position: relative; display: inline-block; }
.dropdown-toggle {
  padding: 8px 12px; font-size: .95rem; border: 1px solid #ccc;
  border-radius: 6px; background: #fff; cursor: pointer;
}
.dropdown-toggle[aria-expanded="true"] { background: #f4f5f7; }
.dropdown-panel {
  position: absolute; z-index: 1000; min-width: 300px; margin-top: 6px;
  background: #fff; border: 1px solid #ddd; border-radius: 8px;
  box-shadow: 0 8px 24px rgba(0,0,0,.12);
  padding: 10px; display: none;
}
.dropdown-panel.open { display: block; }

.dropdown-actions {
  display: flex; gap: 8px; justify-content: flex-end; margin-bottom: 8px;
}
.dropdown-actions button {
  padding: 6px 10px; font-size: .85rem; border: 1px solid #ccc;
  border-radius: 6px; background: #fff; cursor: pointer;
}
.dropdown-actions button:hover { background: #f4f5f7; }

.checkbox-list {
  display: grid; grid-template-columns: 1fr; gap: 6px;
  max-height: 280px; overflow-y: auto; border: 1px solid #eee;
  padding: 8px; border-radius: 6px; background: #fff;
}
.chk { display: flex; align-items: center; gap: 8px; font-size: .95rem; }
/* ▲ ここまで変更（縦並び＆上部アクション） */
"""

# --- 動作（JavaScript） ---
js = r"""
const DATA = __DATA__;
const CHOICES = __CHOICES__;
const COLS = __COLS__;

const exportBtn = document.getElementById('export');

// 年度
const ddYearBtn   = document.getElementById('dd-year-btn');
const ddYearPanel = document.getElementById('dd-year-panel');
const yearList    = document.getElementById('year_list');
const yearSelectAllBtn = document.getElementById('year_select_all');
const yearClearAllBtn  = document.getElementById('year_clear_all');

// 診療科
const ddDeptBtn   = document.getElementById('dd-dept-btn');
const ddDeptPanel = document.getElementById('dd-dept-panel');
const deptList    = document.getElementById('dept_list');
const deptSelectAllBtn = document.getElementById('dept_select_all');
const deptClearAllBtn  = document.getElementById('dept_clear_all');

// 候補描画（縦並び）
function renderYearChoices() {
  yearList.innerHTML = CHOICES['年度']
    .map(y => `<label class="chk"><input type="checkbox" name="year" value="${y}">${y}</label>`)
    .join('');
}
function renderDeptChoices() {
  deptList.innerHTML = CHOICES['診療科']
    .map(d => `<label class="chk"><input type="checkbox" name="dept" value="${d}">${d}</label>`)
    .join('');
}

// 選択値取得（name基準）
function getChecked(name) {
  return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(el => el.value);
}

// フィルタ
function getFiltered() {
  const years = getChecked('year');
  const depts = getChecked('dept');
  return DATA.filter(r =>
    (years.length === 0 || years.includes(r['年度'])) &&
    (depts.length === 0 || depts.includes(r['診療科']))
  );
}

// テーブル生成
function makeTable(containerId, rows) {
  const wrap = document.getElementById(containerId);
  wrap.innerHTML = '';
  if (!rows || rows.length === 0) {
    wrap.innerHTML = '<div class="empty">該当するデータがありません。</div>';
    return;
  }
  const thead = '<thead><tr>' + COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead>';
  const tbody = '<tbody>' + rows.map(r => '<tr>' + COLS.map(c => `<td>${(r[c] ?? '')}</td>`).join('') + '</tr>').join('') + '</tbody>';
  const html = '<div class="tablewrap"><table>' + thead + tbody + '</table></div>';
  wrap.innerHTML = html;
}

// バッジ更新
function renderBadges() {
  const years = getChecked('year');
  const depts = getChecked('dept');
  document.getElementById('badge_year').textContent = years.length ? years.join(', ') : '未選択';
  document.getElementById('badge_dept').textContent = depts.length ? depts.join(', ') : '未選択';
}

// 再描画フロー
function renderCards() {
  const filtered = getFiltered();
  document.getElementById('count').textContent = `${filtered.length} 件`;
  makeTable('tbl_main', filtered);
  renderBadges();
}

// CSV 出力
function exportCSV() {
  const filtered = getFiltered();
  if (filtered.length === 0) { alert('出力対象がありません。'); return; }
  const cols = COLS;
  const rows = [cols.join(',')].concat(filtered.map(r =>
    cols.map(c => String(r[c] ?? '').replaceAll('"', '""'))
        .map(v => /[",\n]/.test(v) ? `"${v}"` : v)
        .join(',')
  ));
  const blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = '地域貢献_filtered_export.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ドロップダウン開閉
function toggleDropdown(btn, panel, open) {
  const isOpen = (open != null) ? open : !(panel.classList.contains('open'));
  panel.classList.toggle('open', isOpen);
  btn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
}
ddYearBtn.addEventListener('click', () => toggleDropdown(ddYearBtn, ddYearPanel));
ddDeptBtn.addEventListener('click', () => toggleDropdown(ddDeptBtn, ddDeptPanel));
document.addEventListener('click', (e) => {
  if (!ddYearBtn.contains(e.target) && !ddYearPanel.contains(e.target)) toggleDropdown(ddYearBtn, ddYearPanel, false);
  if (!ddDeptBtn.contains(e.target) && !ddDeptPanel.contains(e.target)) toggleDropdown(ddDeptBtn, ddDeptPanel, false);
});
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    toggleDropdown(ddYearBtn, ddYearPanel, false);
    toggleDropdown(ddDeptBtn, ddDeptPanel, false);
  }
});

// 一括選択/解除（上部）
yearSelectAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="year"]').forEach(el => el.checked = true);
  renderCards();
});
yearClearAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="year"]').forEach(el => el.checked = false);
  renderCards();
});
deptSelectAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="dept"]').forEach(el => el.checked = true);
  renderCards();
});
deptClearAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="dept"]').forEach(el => el.checked = false);
  renderCards();
});

// 変更即時反映
document.addEventListener('change', (e) => {
  if (e.target && (e.target.name === 'year' || e.target.name === 'dept')) renderCards();
});

// 初期描画
renderYearChoices();
renderDeptChoices();
renderCards();
exportBtn.addEventListener('click', exportCSV);
"""

# --- HTMLテンプレート ---
html_tpl = """<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>地域貢献</title>
<style>[[CSS]]</style>
</head>
<body>
  <header>
    <h1>地域貢献</h1>
    <div class="controls" role="region" aria-label="検索条件">
      <!-- 年度（dropdown＋縦並びチェック） -->
      <div class="dropdown">
        <button class="dropdown-toggle" id="dd-year-btn" aria-expanded="false" aria-controls="dd-year-panel">
          年度を選択（複数可）
        </button>
        <div class="dropdown-panel" id="dd-year-panel" role="listbox" aria-labelledby="dd-year-btn">
          <div class="dropdown-actions">
            <button id="year_select_all" type="button">すべて選択</button>
            <button id="year_clear_all" type="button">すべて解除</button>
          </div>
          <div id="year_list" class="checkbox-list" aria-label="年度選択"></div>
        </div>
      </div>

      <!-- 診療科（dropdown＋縦並びチェック） -->
      <div class="dropdown">
        <button class="dropdown-toggle" id="dd-dept-btn" aria-expanded="false" aria-controls="dd-dept-panel">
          診療科を選択（複数可）
        </button>
        <div class="dropdown-panel" id="dd-dept-panel" role="listbox" aria-labelledby="dd-dept-btn">
          <div class="dropdown-actions">
            <button id="dept_select_all" type="button">すべて選択</button>
            <button id="dept_clear_all" type="button">すべて解除</button>
          </div>
          <div id="dept_list" class="checkbox-list" aria-label="診療科選択"></div>
        </div>
      </div>

      <div style="align-self: end;">
        <button id="export" title="現在の抽出結果をCSVで保存">CSVダウンロード</button>
        <div class="note">※年度・診療科は複数選択できます（未選択の場合は全件）</div>
      </div>
    </div>

    <div class="note">選択中 → 年度: <span class="badge" id="badge_year"></span> ／ 診療科: <span class="badge" id="badge_dept"></span></div>
  </header>

  <section class="card">
    <h2>地域貢献－検索結果</h2>
    <div class="count" id="count"></div>
    <div id="tbl_main"></div>
  </section>

  <script>[[JS]]</script>
</body>
</html>"""

# --- HTML生成 ---
html = (
    html_tpl.replace("[[CSS]]", css)
            .replace("[[SRC]]", os.path.basename(EXCEL_FILE))
            .replace("[[TS]]", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            .replace("[[SHEET]]", str(sheet_name))
)
js_filled = (
    js.replace("__DATA__", json.dumps(records, ensure_ascii=False))
      .replace("__CHOICES__", json.dumps(choices, ensure_ascii=False))
      .replace("__COLS__", json.dumps(cols, ensure_ascii=False))
)
html = html.replace("[[JS]]", js_filled)

with open(HTML_FILE, "w", encoding="utf-8") as f:
    f.write(html)
print("✓ 生成完了:", HTML_FILE)
