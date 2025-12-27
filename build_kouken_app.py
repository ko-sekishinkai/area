
# -*- coding: utf-8 -*-
# 地域貢献データ 年度×診療科 検索アプリ 再生成スクリプト
# 実行: python build_kouken_app.py
# 依存: pandas (openpyxl)

import pandas as pd
import json
from datetime import datetime
import os, re

EXCEL_FILE = "地域貢献_統合.xlsx"
HTML_FILE  = "地域貢献_年度診療科検索_app.html"

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

# --- 選択肢生成（年度 → 診療科）---
if ("年度" not in df.columns) or ("診療科" not in df.columns):
    raise ValueError("Excelに『年度』『診療科』列が必要です。")

years   = sorted(list({y for y in df["年度"].tolist() if y}))
choices = {"年度": years, "診療科By年度": {}}
for y in years:
    depts = sorted(list({d for (yy, d) in df.loc[df["年度"] == y, ["年度", "診療科"]].values if d}))
    choices["診療科By年度"][y] = depts

# --- 表示列の並び（主要項目を前に） ---
preferred = ["年度","診療科","日付","事業所","発表者","タイトル","主催/共催","形態","特記事項（年代、エリア限定等）"]
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
select { padding: 8px 10px; font-size: .95rem; }
.card { border: 1px solid #ddd; border-radius: 8px; padding: 12px; margin: 12px 0; }
.card h2 { font-size: 1.2rem; margin: 0 0 8px; }
.count { color: #333; font-size: .95rem; margin-bottom: 8px; }
.tablewrap { overflow-x: auto; border: 1px solid #eee; border-radius: 6px; }
table { border-collapse: collapse; width: 100%; min-width: 960px; }
th, td { padding: 8px 10px; border-bottom: 1px solid #eee; text-align: left; }
th { background: #f8f9fb; position: sticky; top: 0; z-index: 1; }
tr:nth-child(even) td { background: #fcfcff; }
.empty { color: #666; padding: 12px; }
.footer { margin-top: 18px; color: #555; font-size: .9rem; }
button { padding: 8px 12px; font-size: .9rem; border: 1px solid #ccc; border-radius: 6px; cursor: pointer; background: #fff; }
button:hover { background: #f4f5f7; }
.note { color: #777; font-size: .85rem; }
.badge { display:inline-block; padding: 2px 8px; background: #eef2ff; border:1px solid #c7d2fe; border-radius: 999px; font-size: .8rem; color:#1e40af; }
"""

# --- 動作（JavaScript） ---
js = r"""
const DATA = __DATA__;
const CHOICES = __CHOICES__;
const COLS = __COLS__;
const yearSel = document.getElementById('year');
const deptSel = document.getElementById('dept');
const exportBtn = document.getElementById('export');

function renderYearChoices() {
  yearSel.innerHTML = '<option value=\"\">年度を選択…</option>' + CHOICES['年度'].map(y => `<option value=\"${y}\">${y}</option>`).join('');
  deptSel.innerHTML = '<option value=\"\">診療科を選択…</option>';
}
function updateDeptChoices() {
  const y = yearSel.value;
  const list = (CHOICES['診療科By年度'][y] || []);
  deptSel.innerHTML = '<option value=\"\">診療科を選択…</option>' + list.map(d => `<option value=\"${d}\">${d}</option>`).join('');
}
function getFiltered() {
  const y = yearSel.value;
  const d = deptSel.value;
  return DATA.filter(r => (!y || r['年度'] === y) && (!d || r['診療科'] === d));
}
function makeTable(containerId, rows) {
  const wrap = document.getElementById(containerId);
  wrap.innerHTML = '';
  if (!rows || rows.length === 0) {
    wrap.innerHTML = '<div class=\"empty\">該当するデータがありません。</div>';
    return;
  }
  const thead = '<thead><tr>' + COLS.map(c => `<th>${c}</th>`).join('') + '</tr></thead>';
  const tbody = '<tbody>' + rows.map(r => '<tr>' + COLS.map(c => `<td>${(r[c] ?? '')}</td>`).join('') + '</tr>').join('') + '</tbody>';
  const html = '<div class=\"tablewrap\"><table>' + thead + tbody + '</table></div>';
  wrap.innerHTML = html;
}
function renderCards() {
  const filtered = getFiltered();
  document.getElementById('count').textContent = `${filtered.length} 件`;
  makeTable('tbl_main', filtered);
  const badgeY = document.getElementById('badge_year');
  const badgeD = document.getElementById('badge_dept');
  badgeY.textContent = (yearSel.value ? yearSel.value : '未選択');
  badgeD.textContent = (deptSel.value ? deptSel.value : '未選択');
}
function exportCSV() {
  const filtered = getFiltered();
  if (filtered.length === 0) { alert('出力対象がありません。'); return; }
  const cols = COLS;
  const rows = [cols.join(',')].concat(filtered.map(r =>
    cols.map(c => String(r[c] ?? '').replaceAll('\"', '\"\"'))
        .map(v => /[\",\\n]/.test(v) ? `\"${v}\"` : v)
        .join(',')
  ));
  const blob = new Blob([rows.join('\\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = '地域貢献_filtered_export.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}
yearSel.addEventListener('change', () => { updateDeptChoices(); renderCards(); });
deptSel.addEventListener('change', () => { renderCards(); });
exportBtn.addEventListener('click', exportCSV);
renderYearChoices(); renderCards();
"""

# --- HTMLテンプレート ---
html_tpl = """<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>地域貢献データ 年度×診療科 検索アプリ</title>
<style>[[CSS]]</style>
</head>
<body>
  <header>
    <h1>地域貢献データ 年度×診療科 検索アプリ</h1>
    <div class="meta">ソース: [[SRC]] ／ 生成日時: [[TS]] ／ シート: [[SHEET]]</div>
    <div class="controls">
      <div>
        <label for="year">年度</label><br>
        <select id="year" aria-label="年度選択"></select>
      </div>
      <div>
        <label for="dept">診療科</label><br>
        <select id="dept" aria-label="診療科選択"></select>
      </div>
      <div style="align-self: end;">
        <button id="export" title="現在の抽出結果をCSVで保存">CSVダウンロード</button>
        <div class="note">※年度を選ぶと診療科の選択肢が絞り込まれます。</div>
      </div>
    </div>
    <div class="note">選択中 → 年度: <span class="badge" id="badge_year"></span> ／ 診療科: <span class="badge" id="badge_dept"></span></div>
  </header>

  <section class="card">
    <h2>抽出結果</h2>
    <div class="count" id="count"></div>
    <div id="tbl_main"></div>
  </section>

  <div class="footer">このページはExcelから自動生成されています。毎年Excelを更新したら、Pythonで再生成してください。</div>
  <script>[[JS]]</script>
</body>
</html>"""

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
