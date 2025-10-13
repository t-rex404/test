# -*- coding: utf-8 -*-
from flask import Flask, jsonify
import datetime
import json
import os

app = Flask(__name__)

# JSONディレクトリパス
JSON_DIR = os.path.join(os.path.dirname(__file__), '_json')

# データ読み込み関数
def load_json(filename):
    """JSONファイルを読み込む"""
    filepath = os.path.join(JSON_DIR, filename)
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"⚠️  警告: {filename} が見つかりません")
        return None
    except json.JSONDecodeError as e:
        print(f"⚠️  エラー: {filename} の読み込みに失敗しました - {e}")
        return None

# データを読み込み
areas_data = load_json('areas.json') or []
countries_data = load_json('countries.json') or []
coffees_data = load_json('coffees.json') or []
roast_levels_data = load_json('roast_levels.json') or {"groups": [], "levels": []}
roast_modifiers_data = load_json('roast_modifiers.json') or {}
enums_data = load_json('enums.json') or {}

# 味のラベル定義
taste_labels = {
    "acidity": "酸味",
    "bitterness": "苦味",
    "aroma": "香り",
    "body": "コク",
    "sweetness": "甘さ"
}

# 焙煎レベルのグループ名マッピング
roast_group_mapping = {lvl['id']: lvl['group_id'] for lvl in roast_levels_data.get('levels', [])}
roast_group_names = {grp['id']: grp['name'] for grp in roast_levels_data.get('groups', [])}

def get_country_name(country_id):
    """国IDから国名を取得"""
    country = next((c for c in countries_data if c['id'] == country_id), None)
    return country['name'] if country else country_id

def get_roast_group_name(roast_level_id):
    """焙煎レベルIDからグループ名を取得"""
    group_id = roast_group_mapping.get(roast_level_id, 'medium')
    return roast_group_names.get(group_id, '中煎り')

def calculate_display_taste(base_taste, roast_level_id, modifiers):
    """焙煎による味の変化を計算（base_tasteに存在する属性のみ、表示順序を保証）"""
    modifier = modifiers.get(roast_level_id, {})
    # 表示順序を定義: 酸味、苦味、香り、コク、甘さ
    display_order = ['acidity', 'bitterness', 'aroma', 'body', 'sweetness']

    # OrderedDictを使用して順序を保証
    from collections import OrderedDict
    result = OrderedDict()

    # 定義された順序でbase_tasteに存在する属性のみを計算
    for axis in display_order:
        if axis in base_taste:
            base_value = base_taste.get(axis, 3)
            mod_value = modifier.get(axis, 0)
            # 1-5の範囲にクランプ
            result[axis] = max(1, min(5, base_value + mod_value))

    return result

def get_available_roast_levels(coffee):
    """利用可能な焙煎レベルを取得（コーヒーに指定があればそれを、なければ全て）"""
    available = coffee.get('available_roast_level_ids', [])
    if available:
        return [lvl for lvl in roast_levels_data.get('levels', []) if lvl['id'] in available]
    # デフォルトは全ての焙煎レベル
    return roast_levels_data.get('levels', [])

def create_product_from_coffee(coffee, roast_level_id=None):
    """JSONデータから商品オブジェクトを生成"""
    # デフォルトの焙煎レベルを使用
    if not roast_level_id:
        roast_level_id = f"{coffee.get('roast_level', 'medium')}-medium"

    # 焙煎レベル情報を取得
    roast_level = next((lvl for lvl in roast_levels_data.get('levels', [])
                       if lvl['id'] == roast_level_id), None)
    if not roast_level:
        # フォールバック
        roast_level = {'id': 'medium-medium', 'name': 'ミディアムロースト'}
        roast_level_id = 'medium-medium'

    # 味を計算
    base_taste = coffee.get('taste', {})
    display_taste = calculate_display_taste(base_taste, roast_level_id, roast_modifiers_data)

    # フレーバーノートの抽出（味の特徴から）
    flavor_notes = []
    if display_taste.get('acidity', 0) >= 4:
        flavor_notes.append('フルーティー')
    if display_taste.get('sweetness', 0) >= 4:
        flavor_notes.append('甘み')
    if display_taste.get('bitterness', 0) >= 4:
        flavor_notes.append('ビター')
    if display_taste.get('body', 0) >= 4:
        flavor_notes.append('コク深い')
    if display_taste.get('aroma', 0) >= 4:
        flavor_notes.append('芳香')
    if not flavor_notes:
        flavor_notes = ['バランス']

    # 価格を取得（現在は200gのみ対応）
    price_jpy = coffee.get('price_jpy', {})
    # 200gの価格を優先、なければ最初に見つかった価格を使用
    if '200' in price_jpy:
        default_price = price_jpy['200']
        default_weight = "200g"
    else:
        # フォールバック: 最初に見つかった重量と価格を使用
        first_weight = next(iter(price_jpy.keys()), '200')
        default_price = price_jpy.get(first_weight, 1500)
        default_weight = f"{first_weight}g"

    # 利用可能なグラム数（将来の拡張用）
    available_grams = coffee.get('availability', {}).get('grams', [200])

    # 各グラム数の価格情報を保持（将来の重量選択機能用）
    price_options = []
    for gram in available_grams:
        gram_str = str(gram)
        if gram_str in price_jpy:
            price_options.append({
                'grams': gram,
                'price': price_jpy[gram_str]
            })

    # アイコンを生成（国ごとに）
    icon_map = {
        'br': '☕',  # ブラジル
        'co': '🌰',  # コロンビア
        'et': '🌸',  # エチオピア
        'gt': '🏔️',  # グアテマラ
        'ke': '🍇',  # ケニア
        'id': '🌿'   # インドネシア
    }
    country_id = coffee.get('country_id', 'br')
    icon = icon_map.get(country_id, '☕')

    # 利用可能な焙煎レベルを取得
    available_roast_levels = get_available_roast_levels(coffee)

    return {
        'id': coffee['id'],
        'coffee_id': coffee['id'],  # 元のコーヒーID
        'name': coffee['name'],
        'origin': get_country_name(coffee.get('country_id', '')),
        'roast_level': get_roast_group_name(roast_level_id),
        'roast_level_detail': roast_level['name'],
        'roast_level_id': roast_level_id,
        'available_roast_levels': [{'id': lvl['id'], 'name': lvl['name'], 'group_id': lvl['group_id']}
                                   for lvl in available_roast_levels],
        'price': default_price,
        'weight': default_weight,
        'price_options': price_options,  # 将来の重量選択機能用
        'description': coffee.get('description', ''),
        'flavor_notes': flavor_notes,
        'taste': display_taste,
        'image': icon,
        'varietal': coffee.get('varietal', []),
        'process': coffee.get('process', ''),
        'availability': coffee.get('availability', {}).get('status', 'in_stock')
    }

# 全商品リストを生成
def generate_all_products():
    """全てのコーヒー豆と焙煎レベルの組み合わせから商品を生成"""
    products = []
    product_id = 1

    for coffee in coffees_data:
        # デフォルトの焙煎レベルで商品を作成
        default_roast = f"{coffee.get('roast_level', 'medium')}-medium"
        product = create_product_from_coffee(coffee, default_roast)
        product['id'] = product_id
        products.append(product)
        product_id += 1

    return products

# 商品リストを生成
coffee_products = generate_all_products()

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Café Aroma - コーヒー豆専門店</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', 'Hiragino Sans', 'Yu Gothic', sans-serif;
            background: linear-gradient(135deg, #f5f0e8 0%, #e8dcc8 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .header {
            text-align: center;
            padding: 40px 20px;
            background: linear-gradient(135deg, #6B4423 0%, #8B5A3C 100%);
            color: white;
            border-radius: 15px;
            margin-bottom: 40px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }

        .header h1 {
            font-size: 42px;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }

        .header p {
            font-size: 18px;
            opacity: 0.95;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
        }

        .filter-section {
            background: white;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 30px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            text-align: center;
        }

        .filter-section h3 {
            margin-bottom: 15px;
            color: #6B4423;
        }

        .filter-buttons {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            justify-content: center;
        }

        .filter-btn {
            padding: 10px 20px;
            border: 2px solid #8B5A3C;
            background: white;
            color: #6B4423;
            border-radius: 25px;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.3s ease;
        }

        .filter-btn:hover,
        .filter-btn.active {
            background: #8B5A3C;
            color: white;
            transform: translateY(-2px);
            box-shadow: 0 4px 10px rgba(139,90,60,0.3);
        }

        .products-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 30px;
            margin-bottom: 40px;
        }

        .product-card {
            background: white;
            border-radius: 15px;
            overflow: hidden;
            box-shadow: 0 8px 25px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            cursor: pointer;
        }

        .product-card:hover {
            transform: translateY(-8px);
            box-shadow: 0 15px 40px rgba(0,0,0,0.2);
        }

        .product-icon {
            font-size: 80px;
            text-align: center;
            padding: 30px;
            background: linear-gradient(135deg, #f8f5f0 0%, #f0e8d8 100%);
        }

        .product-info {
            padding: 25px;
        }

        .product-header {
            display: flex;
            justify-content: space-between;
            align-items: start;
            margin-bottom: 15px;
        }

        .product-name {
            font-size: 22px;
            font-weight: bold;
            color: #6B4423;
            margin-bottom: 5px;
        }

        .product-origin {
            color: #8B5A3C;
            font-size: 14px;
            margin-bottom: 3px;
        }

        .product-roast {
            display: inline-block;
            padding: 4px 12px;
            background: #f0e8d8;
            border-radius: 15px;
            font-size: 12px;
            color: #6B4423;
            margin-top: 5px;
        }

        .product-price {
            font-size: 28px;
            font-weight: bold;
            color: #D4A574;
            white-space: nowrap;
        }

        .product-weight {
            font-size: 12px;
            color: #999;
        }

        .product-description {
            color: #666;
            font-size: 14px;
            line-height: 1.6;
            margin: 15px 0;
        }

        .taste-profile {
            margin: 15px 0;
            padding: 15px;
            background: #f9f9f9;
            border-radius: 8px;
        }

        .taste-profile h4 {
            font-size: 13px;
            color: #6B4423;
            margin-bottom: 10px;
            font-weight: 600;
        }

        .taste-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin: 8px 0;
            font-size: 13px;
        }

        .taste-label {
            color: #666;
            min-width: 60px;
        }

        .taste-cups {
            flex: 1;
            display: flex;
            gap: 3px;
            margin: 0 10px;
        }

        .taste-cups .cup {
            font-size: 14px;
        }

        .taste-cups .cup.filled {
            opacity: 1;
        }

        .taste-cups .cup.empty {
            opacity: 0.2;
        }

        .flavor-notes {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin-top: 15px;
        }

        .flavor-tag {
            background: linear-gradient(135deg, #8B5A3C 0%, #6B4423 100%);
            color: white;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 12px;
        }

        .roast-selector {
            margin: 15px 0;
        }

        .roast-selector label {
            display: block;
            font-size: 13px;
            color: #6B4423;
            font-weight: 600;
            margin-bottom: 8px;
        }

        .roast-selector select {
            width: 100%;
            padding: 10px;
            border: 2px solid #D4A574;
            border-radius: 8px;
            background: white;
            color: #6B4423;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.3s ease;
        }

        .roast-selector select:hover {
            border-color: #8B5A3C;
        }

        .roast-selector select:focus {
            outline: none;
            border-color: #6B4423;
            box-shadow: 0 0 0 3px rgba(107, 68, 35, 0.1);
        }

        .footer {
            text-align: center;
            padding: 30px;
            color: #8B5A3C;
            font-size: 14px;
        }

        .data-source {
            background: #fff;
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
            text-align: center;
            color: #6B4423;
            font-size: 13px;
        }

        @media (max-width: 768px) {
            .products-grid {
                grid-template-columns: 1fr;
            }

            .header h1 {
                font-size: 32px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>☕ Café Aroma</h1>
            <p>世界中から厳選したプレミアムコーヒー豆をお届けします</p>
        </div>

        <div class="data-source">
            📂 商品データは _json ディレクトリから読み込まれています
        </div>

        <div class="filter-section">
            <h3>焙煎度で絞り込み</h3>
            <div class="filter-buttons">
                <button class="filter-btn active" onclick="filterProducts('all')">すべて</button>
                <button class="filter-btn" onclick="filterProducts('浅煎り')">浅煎り</button>
                <button class="filter-btn" onclick="filterProducts('中煎り')">中煎り</button>
                <button class="filter-btn" onclick="filterProducts('深煎り')">深煎り</button>
            </div>
        </div>

        <div class="products-grid" id="productsGrid">
            <!-- Products will be loaded here -->
        </div>

        <div class="footer">
            <p>🌍 世界のコーヒー産地から、香り高い豆をお届け</p>
            <p style="margin-top: 10px;">営業時間: 10:00 - 19:00 | 定休日: 水曜日</p>
        </div>
    </div>

    <script>
        let allProducts = [];
        let currentFilter = 'all';

        // Load products from API
        async function loadProducts() {
            const response = await fetch('/api/products');
            allProducts = await response.json();
            displayProducts(allProducts);
        }

        function getTasteCups(level) {
            const maxLevel = 5;
            let cups = '';
            for (let i = 1; i <= maxLevel; i++) {
                if (i <= level) {
                    cups += '<span class="cup filled">☕</span>';
                } else {
                    cups += '<span class="cup empty">☕</span>';
                }
            }
            return cups;
        }

        async function changeRoastLevel(coffeeId, roastLevelId, cardIndex) {
            try {
                const response = await fetch(`/api/coffee/${coffeeId}/roast/${roastLevelId}`);
                const updatedProduct = await response.json();

                // 商品データを更新
                allProducts[cardIndex] = updatedProduct;

                // 該当する商品カードのみを再描画
                displayProducts(currentFilter === 'all' ? allProducts :
                    allProducts.filter(p => p.roast_level === currentFilter));
            } catch (error) {
                console.error('焙煎レベルの変更に失敗しました:', error);
            }
        }

        function displayProducts(products) {
            const grid = document.getElementById('productsGrid');
            const tasteLabels = {
                'acidity': '酸味',
                'bitterness': '苦味',
                'aroma': '香り',
                'body': 'コク',
                'sweetness': '甘さ'
            };

            // 表示順序を定義: 酸味、苦味、香り、コク、甘さ
            const tasteOrder = ['acidity', 'bitterness', 'aroma', 'body', 'sweetness'];

            grid.innerHTML = products.map((product, index) => {
                const roastOptions = product.available_roast_levels || [];
                const globalIndex = allProducts.findIndex(p => p.coffee_id === product.coffee_id);

                return `
                <div class="product-card" data-roast="${product.roast_level}">
                    <div class="product-icon">${product.image}</div>
                    <div class="product-info">
                        <div class="product-header">
                            <div>
                                <div class="product-name">${product.name}</div>
                                <div class="product-origin">📍 ${product.origin}</div>
                                <span class="product-roast">${product.roast_level}</span>
                            </div>
                            <div style="text-align: right;">
                                <div class="product-price">¥${product.price.toLocaleString()}</div>
                                <div class="product-weight">${product.weight}</div>
                            </div>
                        </div>
                        <div class="product-description">${product.description}</div>

                        ${roastOptions.length > 1 ? `
                        <div class="roast-selector">
                            <label>焙煎度合いを選択:</label>
                            <select onchange="changeRoastLevel('${product.coffee_id}', this.value, ${globalIndex})">
                                ${roastOptions.map(roast => `
                                    <option value="${roast.id}" ${roast.id === product.roast_level_id ? 'selected' : ''}>
                                        ${roast.name}
                                    </option>
                                `).join('')}
                            </select>
                        </div>
                        ` : ''}

                        <div class="taste-profile">
                            <h4>味のプロファイル</h4>
                            ${tasteOrder.filter(key => key in product.taste).map(key => `
                                <div class="taste-row">
                                    <span class="taste-label">${tasteLabels[key] || key}</span>
                                    <div class="taste-cups">${getTasteCups(product.taste[key])}</div>
                                </div>
                            `).join('')}
                        </div>

                        <div class="flavor-notes">
                            ${product.flavor_notes.map(note => `
                                <span class="flavor-tag">${note}</span>
                            `).join('')}
                        </div>
                    </div>
                </div>
            `}).join('');
        }

        function filterProducts(roastLevel) {
            // Update button states
            document.querySelectorAll('.filter-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            event.target.classList.add('active');

            currentFilter = roastLevel;

            if (roastLevel === 'all') {
                displayProducts(allProducts);
            } else {
                const filtered = allProducts.filter(p => p.roast_level === roastLevel);
                displayProducts(filtered);
            }
        }

        // Initial load
        loadProducts();
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    """トップページ - コーヒー豆商品一覧"""
    return HTML_TEMPLATE

@app.route('/api/products', methods=['GET'])
def get_products():
    """商品一覧を取得するAPI"""
    return jsonify(coffee_products)

@app.route('/api/products/<int:product_id>', methods=['GET'])
def get_product(product_id):
    """特定の商品を取得するAPI"""
    product = next((p for p in coffee_products if p["id"] == product_id), None)
    if product:
        return jsonify(product)
    return jsonify({"error": "商品が見つかりません"}), 404

@app.route('/api/coffee/<string:coffee_id>/roast/<string:roast_level_id>', methods=['GET'])
def get_coffee_with_roast(coffee_id, roast_level_id):
    """特定のコーヒー豆と焙煎レベルの組み合わせで味を計算して返すAPI"""
    coffee = next((c for c in coffees_data if c["id"] == coffee_id), None)
    if not coffee:
        return jsonify({"error": "コーヒー豆が見つかりません"}), 404

    roast_level = next((lvl for lvl in roast_levels_data.get('levels', [])
                       if lvl['id'] == roast_level_id), None)
    if not roast_level:
        return jsonify({"error": "焙煎レベルが見つかりません"}), 404

    product = create_product_from_coffee(coffee, roast_level_id)
    return jsonify(product)

@app.route('/api/roast-levels', methods=['GET'])
def get_roast_levels():
    """焙煎レベル一覧を取得するAPI"""
    return jsonify(roast_levels_data)

@app.route('/api/countries', methods=['GET'])
def get_countries():
    """国一覧を取得するAPI"""
    return jsonify(countries_data)

@app.route('/api/info')
def info():
    """システム情報を返すAPI"""
    return jsonify({
        "app_name": "Café Aroma - コーヒー豆専門店",
        "version": "2.0.0",
        "current_time": datetime.datetime.now().isoformat(),
        "products_count": len(coffee_products),
        "data_source": "_json ディレクトリ"
    })

if __name__ == '__main__':
    print("=" * 60)
    print("☕ Café Aroma - コーヒー豆専門店 Webサービスを起動中...")
    print("=" * 60)
    print(f"📂 データソース: {JSON_DIR}")
    print(f"📊 読み込まれた商品数: {len(coffee_products)}")
    print(f"🌍 対応国数: {len(countries_data)}")
    print(f"🔥 焙煎レベル数: {len(roast_levels_data.get('levels', []))}")
    print("=" * 60)
    print("📍 アクセス先: http://localhost:5000")
    print("📍 商品API: http://localhost:5000/api/products")
    print("📍 焙煎レベルAPI: http://localhost:5000/api/roast-levels")
    print("📍 システム情報: http://localhost:5000/api/info")
    print("📍 終了するには: Ctrl+C")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=5000)
