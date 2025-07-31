// 検索機能の実装
class SearchManager {
    constructor() {
        this.searchBox = document.getElementById('searchBox');
        this.searchResults = document.getElementById('searchResults');
        this.searchContainer = document.getElementById('searchContainer');
        this.allContent = [];
        
        this.init();
    }
    
    init() {
        if (this.searchBox) {
            this.searchBox.addEventListener('input', (e) => this.handleSearch(e.target.value));
            this.searchBox.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    this.performSearch(e.target.value);
                }
            });
        }
        
        // ページ内の検索対象コンテンツを収集
        this.collectSearchableContent();
    }
    
    collectSearchableContent() {
        // 検索対象となる要素を収集
        const searchableElements = document.querySelectorAll('h1, h2, h3, h4, p, li, code, pre');
        
        searchableElements.forEach((element, index) => {
            const text = element.textContent.trim();
            if (text.length > 3) { // 3文字以上のテキストのみ対象
                this.allContent.push({
                    element: element,
                    text: text.toLowerCase(),
                    originalText: text,
                    type: element.tagName.toLowerCase(),
                    index: index
                });
            }
        });
    }
    
    handleSearch(query) {
        if (query.length < 2) {
            this.hideSearchResults();
            return;
        }
        
        const results = this.searchContent(query);
        this.displaySearchResults(results, query);
    }
    
    searchContent(query) {
        const lowerQuery = query.toLowerCase();
        const results = [];
        
        this.allContent.forEach(item => {
            if (item.text.includes(lowerQuery)) {
                const relevance = this.calculateRelevance(item.text, lowerQuery);
                results.push({
                    ...item,
                    relevance: relevance,
                    highlightedText: this.highlightMatch(item.originalText, query)
                });
            }
        });
        
        // 関連性でソート
        return results.sort((a, b) => b.relevance - a.relevance).slice(0, 10);
    }
    
    calculateRelevance(text, query) {
        let relevance = 0;
        
        // 完全一致
        if (text === query) relevance += 100;
        
        // 単語の開始位置での一致
        const words = text.split(' ');
        words.forEach(word => {
            if (word.toLowerCase().startsWith(query)) {
                relevance += 50;
            }
        });
        
        // タイトル要素の重み付け
        if (text.includes(query)) {
            relevance += 10;
        }
        
        return relevance;
    }
    
    highlightMatch(text, query) {
        const regex = new RegExp(`(${query})`, 'gi');
        return text.replace(regex, '<mark>$1</mark>');
    }
    
    displaySearchResults(results, query) {
        if (!this.searchResults) return;
        
        if (results.length === 0) {
            this.searchResults.innerHTML = `
                <div class="search-result-item">
                    <p>「${query}」に一致する結果が見つかりませんでした。</p>
                </div>
            `;
        } else {
            const resultsHtml = results.map(result => `
                <div class="search-result-item" onclick="scrollToElement('${result.element.id || 'element-' + result.index}')">
                    <div class="result-type">${this.getTypeLabel(result.type)}</div>
                    <div class="result-content">${result.highlightedText}</div>
                </div>
            `).join('');
            
            this.searchResults.innerHTML = resultsHtml;
        }
        
        this.showSearchResults();
    }
    
    getTypeLabel(type) {
        const labels = {
            'h1': 'タイトル',
            'h2': '見出し',
            'h3': '小見出し',
            'h4': '小見出し',
            'p': '段落',
            'li': 'リスト',
            'code': 'コード',
            'pre': 'コードブロック'
        };
        return labels[type] || 'テキスト';
    }
    
    showSearchResults() {
        if (this.searchResults) {
            this.searchResults.style.display = 'block';
        }
    }
    
    hideSearchResults() {
        if (this.searchResults) {
            this.searchResults.style.display = 'none';
        }
    }
    
    performSearch(query) {
        if (query.trim()) {
            this.handleSearch(query);
        }
    }
}

// 要素にスクロールする関数
function scrollToElement(elementId) {
    const element = document.getElementById(elementId);
    if (element) {
        element.scrollIntoView({
            behavior: 'smooth',
            block: 'center'
        });
        
        // ハイライト効果
        element.style.backgroundColor = '#fff3cd';
        setTimeout(() => {
            element.style.backgroundColor = '';
        }, 2000);
    }
}

// ページ読み込み完了時に検索機能を初期化
document.addEventListener('DOMContentLoaded', function() {
    new SearchManager();
    
    // 検索ボックス外をクリックした時に検索結果を非表示
    document.addEventListener('click', function(e) {
        const searchContainer = document.getElementById('searchContainer');
        const searchResults = document.getElementById('searchResults');
        
        if (searchContainer && searchResults && !searchContainer.contains(e.target)) {
            searchResults.style.display = 'none';
        }
    });
});

// キーボードショートカット
document.addEventListener('keydown', function(e) {
    // Ctrl + K で検索ボックスにフォーカス
    if (e.ctrlKey && e.key === 'k') {
        e.preventDefault();
        const searchBox = document.getElementById('searchBox');
        if (searchBox) {
            searchBox.focus();
        }
    }
    
    // Escape で検索結果を非表示
    if (e.key === 'Escape') {
        const searchResults = document.getElementById('searchResults');
        if (searchResults) {
            searchResults.style.display = 'none';
        }
    }
});

// 検索結果のスタイル
const searchStyles = `
<style>
.search-result-item {
    padding: 1rem;
    border-bottom: 1px solid #e0e0e0;
    cursor: pointer;
    transition: background-color 0.2s ease;
}

.search-result-item:hover {
    background-color: #f8f9fa;
}

.search-result-item:last-child {
    border-bottom: none;
}

.result-type {
    font-size: 0.8rem;
    color: #6c757d;
    font-weight: 500;
    margin-bottom: 0.5rem;
}

.result-content {
    color: #333;
    line-height: 1.4;
}

mark {
    background-color: #fff3cd;
    color: #856404;
    padding: 0.1rem 0.2rem;
    border-radius: 2px;
}

#searchResults {
    position: absolute;
    top: 100%;
    left: 0;
    right: 0;
    background: white;
    border: 1px solid #e0e0e0;
    border-top: none;
    border-radius: 0 0 10px 10px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
    max-height: 400px;
    overflow-y: auto;
    z-index: 1000;
    display: none;
}

.search-container {
    position: relative;
}
</style>
`;

// スタイルをページに追加
document.head.insertAdjacentHTML('beforeend', searchStyles); 