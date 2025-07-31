// PowerShell _lib ライブラリ ドキュメント用JavaScript

// DOM読み込み完了時の初期化
document.addEventListener('DOMContentLoaded', function() {
    initializeSearch();
    initializeNavigation();
    initializeCodeBlocks();
    initializeTables();
});

// 検索機能の初期化
function initializeSearch() {
    const searchBox = document.getElementById('searchBox');
    const searchResults = document.getElementById('searchResults');
    
    if (!searchBox) return;
    
    // Ctrl+K で検索ボックスにフォーカス
    document.addEventListener('keydown', function(e) {
        if (e.ctrlKey && e.key === 'k') {
            e.preventDefault();
            searchBox.focus();
        }
    });
    
    // 検索機能の実装
    searchBox.addEventListener('input', function() {
        const query = this.value.toLowerCase();
        if (query.length > 0) {
            performSearch(query, searchResults);
        } else {
            searchResults.innerHTML = '';
        }
    });
    
    // Enterキーで検索実行
    searchBox.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            const query = this.value.toLowerCase();
            if (query.length > 0) {
                performSearch(query, searchResults);
            }
        }
    });
}

// 検索実行
function performSearch(query, searchResults) {
    // 検索対象の要素を取得
    const searchableElements = document.querySelectorAll('h2, h3, h4, p, li, .code-block pre');
    const results = [];
    
    searchableElements.forEach(element => {
        const text = element.textContent.toLowerCase();
        if (text.includes(query)) {
            const parentSection = element.closest('.detail-page, .card');
            if (parentSection) {
                const title = parentSection.querySelector('h2, h3, h4')?.textContent || '該当箇所';
                const snippet = element.textContent.substring(0, 100) + '...';
                
                results.push({
                    title: title,
                    snippet: snippet,
                    element: element
                });
            }
        }
    });
    
    // 検索結果を表示
    displaySearchResults(results, query, searchResults);
}

// 検索結果の表示
function displaySearchResults(results, query, searchResults) {
    if (results.length === 0) {
        searchResults.innerHTML = `<p>「${query}」に一致する結果が見つかりませんでした。</p>`;
        return;
    }
    
    let html = `<div class="search-results">`;
    html += `<h3>「${query}」の検索結果 (${results.length}件)</h3>`;
    
    results.slice(0, 10).forEach(result => {
        html += `
            <div class="search-result-item">
                <h4>${result.title}</h4>
                <p>${result.snippet}</p>
            </div>
        `;
    });
    
    if (results.length > 10) {
        html += `<p>他に ${results.length - 10} 件の結果があります。</p>`;
    }
    
    html += `</div>`;
    searchResults.innerHTML = html;
}

// ナビゲーション機能の初期化
function initializeNavigation() {
    // ページ内リンクのスムーズスクロール
    const internalLinks = document.querySelectorAll('a[href^="#"]');
    internalLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            const targetId = this.getAttribute('href').substring(1);
            const targetElement = document.getElementById(targetId);
            if (targetElement) {
                targetElement.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });
    
    // 戻るリンクの追加
    addBackLinks();
}

// 戻るリンクの追加
function addBackLinks() {
    const detailPages = document.querySelectorAll('.detail-page');
    detailPages.forEach(page => {
        const firstHeading = page.querySelector('h2, h3');
        if (firstHeading && !page.querySelector('.back-link')) {
            const backLink = document.createElement('a');
            backLink.href = 'index.html';
            backLink.className = 'back-link';
            backLink.textContent = '← トップページに戻る';
            page.insertBefore(backLink, firstHeading);
        }
    });
}

// コードブロックの初期化
function initializeCodeBlocks() {
    const codeBlocks = document.querySelectorAll('.code-block pre');
    codeBlocks.forEach(block => {
        // コードブロックにコピーボタンを追加
        addCopyButton(block);
        
        // シンタックスハイライト（簡易版）
        highlightSyntax(block);
    });
}

// コピーボタンの追加
function addCopyButton(codeBlock) {
    const copyButton = document.createElement('button');
    copyButton.className = 'copy-button';
    copyButton.textContent = 'コピー';
    copyButton.style.cssText = `
        position: absolute;
        top: 5px;
        right: 5px;
        background: #007acc;
        color: white;
        border: none;
        padding: 5px 10px;
        border-radius: 3px;
        cursor: pointer;
        font-size: 12px;
    `;
    
    copyButton.addEventListener('click', function() {
        const text = codeBlock.textContent;
        navigator.clipboard.writeText(text).then(() => {
            this.textContent = 'コピー完了!';
            setTimeout(() => {
                this.textContent = 'コピー';
            }, 2000);
        }).catch(err => {
            console.error('コピーに失敗しました:', err);
            this.textContent = 'コピー失敗';
            setTimeout(() => {
                this.textContent = 'コピー';
            }, 2000);
        });
    });
    
    // コードブロックの親要素を相対位置に設定
    const parent = codeBlock.parentElement;
    parent.style.position = 'relative';
    parent.appendChild(copyButton);
}

// シンタックスハイライト（簡易版）
function highlightSyntax(codeBlock) {
    const text = codeBlock.textContent;
    
    // PowerShellのキーワードをハイライト
    const keywords = [
        'function', 'param', 'begin', 'process', 'end', 'if', 'else', 'foreach',
        'while', 'do', 'until', 'switch', 'try', 'catch', 'finally', 'throw',
        'return', 'break', 'continue', 'new', 'static', 'class', 'enum',
        'interface', 'namespace', 'using', 'import', 'export', 'module',
        'script', 'workflow', 'configuration', 'filter', 'trap'
    ];
    
    let highlightedText = text;
    keywords.forEach(keyword => {
        const regex = new RegExp(`\\b${keyword}\\b`, 'gi');
        highlightedText = highlightedText.replace(regex, `<span class="keyword">${keyword}</span>`);
    });
    
    // コメントをハイライト
    highlightedText = highlightedText.replace(/#.*$/gm, '<span class="comment">$&</span>');
    
    // 文字列をハイライト
    highlightedText = highlightedText.replace(/"([^"]*)"/g, '<span class="string">"$1"</span>');
    
    codeBlock.innerHTML = highlightedText;
}

// テーブルの初期化
function initializeTables() {
    const tables = document.querySelectorAll('table');
    tables.forEach(table => {
        // テーブルにソート機能を追加
        addTableSorting(table);
        
        // テーブルに検索機能を追加
        addTableSearch(table);
    });
}

// テーブルソート機能
function addTableSorting(table) {
    const headers = table.querySelectorAll('th');
    headers.forEach((header, index) => {
        header.style.cursor = 'pointer';
        header.addEventListener('click', function() {
            sortTable(table, index);
        });
        
        // ソートアイコンを追加
        const sortIcon = document.createElement('span');
        sortIcon.textContent = ' ↕';
        sortIcon.style.fontSize = '12px';
        header.appendChild(sortIcon);
    });
}

// テーブルソート実行
function sortTable(table, columnIndex) {
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    
    rows.sort((a, b) => {
        const aValue = a.cells[columnIndex].textContent.trim();
        const bValue = b.cells[columnIndex].textContent.trim();
        
        // 数値の場合は数値として比較
        const aNum = parseFloat(aValue);
        const bNum = parseFloat(bValue);
        
        if (!isNaN(aNum) && !isNaN(bNum)) {
            return aNum - bNum;
        }
        
        // 文字列の場合は辞書順で比較
        return aValue.localeCompare(bValue, 'ja');
    });
    
    // ソートされた行を再配置
    rows.forEach(row => tbody.appendChild(row));
}

// テーブル検索機能
function addTableSearch(table) {
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = 'テーブル内を検索...';
    searchInput.className = 'table-search';
    searchInput.style.cssText = `
        width: 100%;
        padding: 8px;
        margin: 10px 0;
        border: 1px solid #ddd;
        border-radius: 4px;
        font-size: 14px;
    `;
    
    searchInput.addEventListener('input', function() {
        const query = this.value.toLowerCase();
        const rows = table.querySelectorAll('tbody tr');
        
        rows.forEach(row => {
            const text = row.textContent.toLowerCase();
            if (text.includes(query)) {
                row.style.display = '';
            } else {
                row.style.display = 'none';
            }
        });
    });
    
    table.parentNode.insertBefore(searchInput, table);
}

// ページ読み込み時のアニメーション
function initializeAnimations() {
    const elements = document.querySelectorAll('.card, .detail-page');
    
    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.opacity = '1';
                entry.target.style.transform = 'translateY(0)';
            }
        });
    });
    
    elements.forEach(element => {
        element.style.opacity = '0';
        element.style.transform = 'translateY(20px)';
        element.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        observer.observe(element);
    });
}

// ダークモード切り替え機能
function initializeDarkMode() {
    const darkModeToggle = document.createElement('button');
    darkModeToggle.textContent = '🌙';
    darkModeToggle.className = 'dark-mode-toggle';
    darkModeToggle.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #007acc;
        color: white;
        border: none;
        padding: 10px;
        border-radius: 50%;
        cursor: pointer;
        font-size: 18px;
        z-index: 1000;
    `;
    
    darkModeToggle.addEventListener('click', function() {
        document.body.classList.toggle('dark-mode');
        this.textContent = document.body.classList.contains('dark-mode') ? '☀️' : '🌙';
        
        // ダークモードの状態をローカルストレージに保存
        localStorage.setItem('darkMode', document.body.classList.contains('dark-mode'));
    });
    
    // 保存されたダークモード設定を復元
    const savedDarkMode = localStorage.getItem('darkMode') === 'true';
    if (savedDarkMode) {
        document.body.classList.add('dark-mode');
        darkModeToggle.textContent = '☀️';
    }
    
    document.body.appendChild(darkModeToggle);
}

// 初期化関数の呼び出し
document.addEventListener('DOMContentLoaded', function() {
    initializeSearch();
    initializeNavigation();
    initializeCodeBlocks();
    initializeTables();
    initializeAnimations();
    initializeDarkMode();
}); 