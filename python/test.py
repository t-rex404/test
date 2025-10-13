<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>コーヒーメニュー・焙煎シミュレーター（PyScript）</title>
  <!-- PyScript (CDN)。ネット接続が無い場合はローカルに置いて参照してください。-->
  <link rel="stylesheet" href="https://pyscript.net/releases/2024.5.1/core.css" />
  <script type="module" src="https://pyscript.net/releases/2024.5.1/core.js"></script>
  <style>
    :root { --bg:#0b0c10; --card:#12151a; --ink:#e8edf2; --muted:#a8b0ba; --accent:#36c; }
    html,body { background:var(--bg); color:var(--ink); font-family: system-ui, -apple-system, Segoe UI, Roboto, "Hiragino Kaku Gothic ProN", Meiryo, sans-serif; }
    .container { max-width: 1000px; margin: 24px auto; padding: 0 12px; }
    h1 { font-size: 24px; margin: 8px 0 16px; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
    .card { background: var(--card); border: 1px solid #1f2630; border-radius: 14px; padding: 16px; }
    label { font-size: 12px; color: var(--muted); display:block; margin-bottom: 6px; }
    select { width: 100%; padding: 10px; border-radius: 10px; border:1px solid #2a3340; background:#0f131a; color:var(--ink); }
    .taste { display: grid; grid-template-columns: 90px 1fr 40px; gap: 8px; align-items: center; }
    .bar { height: 10px; background: #233044; border-radius: 6px; position: relative; }
    .bar .fill { position:absolute; left:0; top:0; bottom:0; border-radius:6px; background: linear-gradient(90deg, #6aa9ff, #a184ff); }
    .muted { color: var(--muted); font-size: 12px; }
    .pill { display:inline-block; padding:4px 8px; border-radius:999px; background:#172031; border:1px solid #243049; margin-right:6px; font-size:12px; }
    .row { display:flex; gap:8px; flex-wrap:wrap; }
    .desc { white-space: pre-wrap; line-height: 1.7; }
    .foot { margin-top: 18px; font-size: 12px; color: var(--muted); }
  </style>
</head>
<body>
  <div class="container">
    <h1>コーヒーメニュー・焙煎シミュレーター</h1>
    <div class="grid">
      <div class="card">
        <label>エリア → 国 → 銘柄</label>
        <div class="row">
          <select id="areaSelect"></select>
          <select id="countrySelect"></select>
          <select id="coffeeSelect"></select>
        </div>
        <div style="height:12px"></div>
        <label>焙煎レベル</label>
        <div class="row">
          <select id="roastSelect"></select>
        </div>
        <div class="foot">表示味 = 基準味 + 焙煎補正（1〜5にクランプ）</div>
      </div>

      <div class="card">
        <div id="coffeeName" style="font-weight:600; font-size: 18px; margin-bottom:6px;"></div>
        <div class="muted" id="coffeeMeta" style="margin-bottom: 8px;"></div>
        <div class="desc" id="coffeeDesc"></div>
      </div>
    </div>

    <div style="height:16px"></div>

    <div class="card">
      <div class="row" style="justify-content: space-between; align-items:center; margin-bottom:8px;">
        <div class="muted">味プロファイル</div>
        <div id="roastTag" class="pill"></div>
      </div>
      <div id="tasteGrid" class="grid" style="grid-template-columns: 1fr 1fr;">
        <!-- taste bars injected here -->
      </div>
    </div>

    <div class="foot">これは静的プロトタイプです。データはページ内のJSONから読み込まれます。</div>
  </div>

  <!-- データ: 最小サンプル。実運用では分割JSONをfetchして結合してもOK -->
  <script type="application/json" id="areas-json">{
    "areas": [
      {"id":"sa","name":"南米","order":10},
      {"id":"cam","name":"中南米","order":20}
    ]
  }</script>

  <script type="application/json" id="countries-json">{
    "countries": [
      {"id":"br","name":"ブラジル","area_id":"sa","order":100},
      {"id":"co","name":"コロンビア","area_id":"sa","order":110}
    ]
  }</script>

  <script type="application/json" id="coffees-json">{
    "coffees": [
      {"id":"br-bourbon-peaberry","country_id":"br","name":"ブルボンピーベリー","description":"チョコとナッツの甘さ、なめらかな口当たり。","taste_base":{"acidity":2,"bitterness":2,"aroma":4,"body":3,"sweetness":3},"available_roast_level_ids":["light-cinnamon","medium-high","medium-city","dark-fullcity"]},
      {"id":"br-amarelo-bourbon","country_id":"br","name":"アマレロブルボン","description":"黄果実の甘さ、キャラメル、まろやかな口当たり。","taste_base":{"acidity":3,"bitterness":2,"aroma":4,"body":3,"sweetness":3},"available_roast_level_ids":["light-light","medium-medium","medium-city"]},
      {"id":"co-caturra","country_id":"co","name":"カトゥーラ","description":"柑橘のニュアンス、クリーンでバランス良し。","taste_base":{"acidity":4,"bitterness":1,"aroma":4,"body":2,"sweetness":3},"available_roast_level_ids":["light-light","light-cinnamon","medium-medium","dark-french"]}
    ]
  }</script>

  <script type="application/json" id="roast-levels-json">{
    "groups": [
      {"id":"light","name":"浅煎り","order":10},
      {"id":"medium","name":"中煎り","order":20},
      {"id":"dark","name":"深煎り","order":30}
    ],
    "levels": [
      {"id":"light-light","group_id":"light","name":"ライトロースト","order":11},
      {"id":"light-cinnamon","group_id":"light","name":"シナモンロースト","order":12},
      {"id":"medium-medium","group_id":"medium","name":"ミディアムロースト","order":21},
      {"id":"medium-high","group_id":"medium","name":"ハイロースト","order":22},
      {"id":"medium-city","group_id":"medium","name":"シティロースト","order":23},
      {"id":"dark-fullcity","group_id":"dark","name":"フルシティロースト","order":31},
      {"id":"dark-french","group_id":"dark","name":"フレンチロースト","order":32},
      {"id":"dark-italian","group_id":"dark","name":"イタリアンロースト","order":33}
    ]
  }</script>

  <script type="application/json" id="roast-modifiers-json">{
    "modifiers": {
      "light-light":    {"acidity": 2, "bitterness": -2, "aroma": 1, "body": -1, "sweetness": 0},
      "light-cinnamon": {"acidity": 1, "bitterness": -1, "aroma": 1, "body": 0,  "sweetness": 0},
      "medium-medium":  {"acidity": 0, "bitterness": 0,  "aroma": 1, "body": 0,  "sweetness": 1},
      "medium-high":    {"acidity": -1,"bitterness": 1,  "aroma": 0, "body": 1,  "sweetness": 1},
      "medium-city":    {"acidity": -1,"bitterness": 1,  "aroma": 0, "body": 1,  "sweetness": 1},
      "dark-fullcity":  {"acidity": -2,"bitterness": 1,  "aroma": -1,"body": 2,  "sweetness": 0},
      "dark-french":    {"acidity": -2,"bitterness": 2,  "aroma": -1,"body": 2,  "sweetness": -1},
      "dark-italian":   {"acidity": -2,"bitterness": 2,  "aroma": -2,"body": 2,  "sweetness": -1}
    }
  }</script>

  <!-- PyScript: UIロジック -->
  <py-script>
from js import document
import json

# 1–5でクランプ
def clamp(v, lo=1, hi=5):
    return max(lo, min(hi, v))

# JSONノードから辞書を取得

def load_json(id):
    el = document.getElementById(id)
    return json.loads(el.textContent)

areas = load_json("areas-json")["areas"]
countries = load_json("countries-json")["countries"]
coffees = load_json("coffees-json")["coffees"]
roast_levels = load_json("roast-levels-json")
roast_mods = load_json("roast-modifiers-json")["modifiers"]

axes = ["acidity","bitterness","aroma","body","sweetness"]
axis_labels = {
    "acidity":"酸味","bitterness":"苦味","aroma":"香り","body":"コク","sweetness":"甘さ"
}

# DOM参照
sel_area = document.getElementById("areaSelect")
sel_country = document.getElementById("countrySelect")
sel_coffee = document.getElementById("coffeeSelect")
sel_roast = document.getElementById("roastSelect")

d_coffee_name = document.getElementById("coffeeName")
d_coffee_meta = document.getElementById("coffeeMeta")
d_coffee_desc = document.getElementById("coffeeDesc")
d_roast_tag = document.getElementById("roastTag")

taste_grid = document.getElementById("tasteGrid")

# UIユーティリティ

def set_options(select, items, get_value, get_label):
    select.innerHTML = ""
    for it in items:
        opt = document.createElement("option")
        opt.value = get_value(it)
        opt.textContent = get_label(it)
        select.appendChild(opt)

# 初期化：エリア
areas_sorted = sorted(areas, key=lambda a: a.get("order", 0))
set_options(sel_area, areas_sorted, lambda a: a["id"], lambda a: a["name"])

# 依存セレクタ連動

def on_area_change(event=None):
    aid = sel_area.value
    cands = [c for c in countries if c["area_id"] == aid]
    cands = sorted(cands, key=lambda c: c.get("order", 0))
    set_options(sel_country, cands, lambda c: c["id"], lambda c: c["name"])
    on_country_change()


def on_country_change(event=None):
    cid = sel_country.value
    cands = [cf for cf in coffees if cf["country_id"] == cid]
    cands = sorted(cands, key=lambda c: c.get("name", ""))
    set_options(sel_coffee, cands, lambda c: c["id"], lambda c: c["name"])
    on_coffee_change()


def on_coffee_change(event=None):
    # コーヒー詳細
    cid = sel_coffee.value
    coffee = next((c for c in coffees if c["id"] == cid), None)
    if not coffee:
        return

    d_coffee_name.textContent = coffee["name"]
    country = next((c for c in countries if c["id"] == coffee["country_id"]), None)
    d_coffee_meta.textContent = f"国: {country['name']}"
    d_coffee_desc.textContent = coffee.get("description", "")

    # 焙煎候補
    level_items = [lvl for lvl in roast_levels["levels"] if lvl["id"] in coffee.get("available_roast_level_ids", [])]
    level_items = sorted(level_items, key=lambda l: l.get("order", 0))
    set_options(sel_roast, level_items, lambda l: l["id"], lambda l: f"{next(g['name'] for g in roast_levels['groups'] if g['id']==l['group_id'])} / {l['name']}")

    on_roast_change()


def on_roast_change(event=None):
    cid = sel_coffee.value
    coffee = next((c for c in coffees if c["id"] == cid), None)
    level_id = sel_roast.value
    level = next((l for l in roast_levels["levels"] if l["id"] == level_id), None)
    mod = roast_mods.get(level_id, {})

    d_roast_tag.textContent = level["name"] if level else ""

    # 味バーを再描画
    taste_grid.innerHTML = ""

    for ax in axes:
        row = document.createElement("div")
        row.className = "taste"
        lab = document.createElement("div")
        lab.textContent = axis_labels[ax]
        bar = document.createElement("div")
        bar.className = "bar"
        fill = document.createElement("div")
        fill.className = "fill"

        base = coffee["taste_base"].get(ax, 3)
        delta = mod.get(ax, 0)
        val = clamp(base + delta)
        pct = int(val * 20)  # 1..5 -> 20..100
        fill.style.width = f"{pct}%"

        valbox = document.createElement("div")
        valbox.textContent = str(val)
        valbox.className = "muted"

        bar.appendChild(fill)
        row.appendChild(lab)
        row.appendChild(bar)
        row.appendChild(valbox)
        taste_grid.appendChild(row)

# イベント束縛
sel_area.addEventListener('change', on_area_change)
sel_country.addEventListener('change', on_country_change)
sel_coffee.addEventListener('change', on_coffee_change)
sel_roast.addEventListener('change', on_roast_change)

# 初回描画
on_area_change()
  </py-script>
</body>
</html>
