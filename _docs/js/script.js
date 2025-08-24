// ========================================
// ダークモード切り替え機能
// ========================================
const themeToggle = document.getElementById('themeToggle');
const body = document.body;
const themeIcon = document.querySelector('.theme-icon');
const themeText = document.querySelector('.theme-text');

// ローカルストレージからテーマ設定を読み込み
const savedTheme = localStorage.getItem('theme');
if (savedTheme === 'dark') {
    body.classList.add('dark-mode');
    updateThemeButton();
}

// テーマ切り替えボタンのクリックイベント
themeToggle?.addEventListener('click', () => {
    body.classList.toggle('dark-mode');
    updateThemeButton();
    
    // テーマ設定を保存
    const currentTheme = body.classList.contains('dark-mode') ? 'dark' : 'light';
    localStorage.setItem('theme', currentTheme);
    
    // アニメーション効果
    animateThemeChange();
});

// テーマボタンの表示を更新
function updateThemeButton() {
    if (body.classList.contains('dark-mode')) {
        themeIcon.textContent = '☀️';
        themeText.textContent = 'ライトモード';
    } else {
        themeIcon.textContent = '🌙';
        themeText.textContent = 'ダークモード';
    }
}

// テーマ変更時のアニメーション
function animateThemeChange() {
    const cards = document.querySelectorAll('.class-card, .note-card');
    cards.forEach((card, index) => {
        card.style.animation = 'none';
        setTimeout(() => {
            card.style.animation = `fadeIn 0.6s ease-out ${index * 0.05}s`;
        }, 10);
    });
}

// ========================================
// 検索機能
// ========================================
const searchInput = document.getElementById('searchInput');
const classCards = document.querySelectorAll('.class-card');
const noteCards = document.querySelectorAll('.note-card');

// 検索入力のイベントリスナー
searchInput?.addEventListener('input', (e) => {
    const searchTerm = e.target.value.toLowerCase().trim();
    
    // 検索文字列が空の場合は全て表示
    if (searchTerm === '') {
        showAllCards();
        return;
    }
    
    // カードをフィルタリング
    filterCards(searchTerm);
    
    // 検索結果をハイライト
    highlightSearchResults(searchTerm);
});

// 全てのカードを表示
function showAllCards() {
    classCards.forEach(card => {
        card.classList.remove('hidden');
        card.style.opacity = '1';
        removeHighlights(card);
    });
}

// カードをフィルタリング
function filterCards(searchTerm) {
    let visibleCount = 0;
    
    classCards.forEach(card => {
        const searchableText = card.getAttribute('data-searchable') || '';
        const cardText = (card.textContent + ' ' + searchableText).toLowerCase();
        
        if (cardText.includes(searchTerm)) {
            card.classList.remove('hidden');
            card.style.opacity = '1';
            visibleCount++;
            
            // マッチしたカードにアニメーションを追加
            card.style.animation = 'none';
            setTimeout(() => {
                card.style.animation = 'fadeIn 0.6s ease-out';
            }, 10);
        } else {
            card.classList.add('hidden');
            card.style.opacity = '0.3';
        }
    });
    
    // 検索結果が0件の場合のメッセージ表示
    showNoResultsMessage(visibleCount);
}

// ハイライト処理
function highlightSearchResults(searchTerm) {
    classCards.forEach(card => {
        if (!card.classList.contains('hidden')) {
            highlightTextInElement(card, searchTerm);
        }
    });
}

// 要素内のテキストをハイライト
function highlightTextInElement(element, searchTerm) {
    // 既存のハイライトを削除
    removeHighlights(element);
    
    // テキストノードを検索してハイライト
    const walker = document.createTreeWalker(
        element,
        NodeFilter.SHOW_TEXT,
        null,
        false
    );
    
    const textNodes = [];
    let node;
    
    while (node = walker.nextNode()) {
        if (node.nodeValue.toLowerCase().includes(searchTerm)) {
            textNodes.push(node);
        }
    }
    
    textNodes.forEach(textNode => {
        const span = document.createElement('span');
        span.className = 'search-highlight';
        span.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
        span.style.color = 'white';
        span.style.padding = '2px 4px';
        span.style.borderRadius = '4px';
        
        const regex = new RegExp(`(${searchTerm})`, 'gi');
        const parts = textNode.nodeValue.split(regex);
        
        parts.forEach(part => {
            if (part.toLowerCase() === searchTerm) {
                const highlight = span.cloneNode();
                highlight.textContent = part;
                textNode.parentNode.insertBefore(highlight, textNode);
            } else {
                const text = document.createTextNode(part);
                textNode.parentNode.insertBefore(text, textNode);
            }
        });
        
        textNode.parentNode.removeChild(textNode);
    });
}

// ハイライトを削除
function removeHighlights(element) {
    const highlights = element.querySelectorAll('.search-highlight');
    highlights.forEach(highlight => {
        const parent = highlight.parentNode;
        while (highlight.firstChild) {
            parent.insertBefore(highlight.firstChild, highlight);
        }
        parent.removeChild(highlight);
    });
}

// 検索結果が0件の場合のメッセージ
function showNoResultsMessage(count) {
    const existingMessage = document.querySelector('.no-results-message');
    
    if (count === 0) {
        if (!existingMessage) {
            const message = document.createElement('div');
            message.className = 'no-results-message';
            message.style.cssText = `
                text-align: center;
                padding: 40px;
                font-size: 1.2em;
                color: var(--text-secondary);
                background: var(--card-bg);
                border-radius: 15px;
                margin: 20px 0;
                box-shadow: var(--shadow-medium);
            `;
            message.innerHTML = `
                <p style="font-size: 3em; margin-bottom: 20px;">😔</p>
                <p>検索結果が見つかりませんでした</p>
                <p style="font-size: 0.9em; margin-top: 10px;">別のキーワードで検索してみてください</p>
            `;
            
            const classGrid = document.querySelector('.class-grid');
            classGrid.parentNode.insertBefore(message, classGrid.nextSibling);
        }
    } else {
        if (existingMessage) {
            existingMessage.remove();
        }
    }
}

// ========================================
// ページロード時の初期化
// ========================================
document.addEventListener('DOMContentLoaded', () => {
    // 初期アニメーション
    animateOnLoad();
    
    // スムーススクロール
    initSmoothScroll();
    
    // ツールチップ初期化
    initTooltips();
    
    // リンクデバッグ用のコードを追加
    initLinkDebug();
    
    // パーティクル効果（オプション）
    // initParticles();
});

// ページロード時のアニメーション
function animateOnLoad() {
    const elements = document.querySelectorAll('.class-card, .note-card, .error-range-card');
    elements.forEach((element, index) => {
        element.style.opacity = '0';
        element.style.transform = 'translateY(20px)';
        
        setTimeout(() => {
            element.style.transition = 'all 0.6s ease-out';
            element.style.opacity = '1';
            element.style.transform = 'translateY(0)';
        }, index * 100);
    });
}

// スムーススクロールの初期化
function initSmoothScroll() {
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });
}

// ツールチップの初期化
function initTooltips() {
    const tooltipElements = document.querySelectorAll('[data-tooltip]');
    
    tooltipElements.forEach(element => {
        element.addEventListener('mouseenter', (e) => {
            const tooltip = document.createElement('div');
            tooltip.className = 'tooltip';
            tooltip.textContent = e.target.getAttribute('data-tooltip');
            tooltip.style.cssText = `
                position: absolute;
                background: var(--gradient-1);
                color: white;
                padding: 8px 12px;
                border-radius: 8px;
                font-size: 0.9em;
                z-index: 1000;
                pointer-events: none;
                opacity: 0;
                transition: opacity 0.3s ease;
            `;
            
            document.body.appendChild(tooltip);
            
            const rect = e.target.getBoundingClientRect();
            tooltip.style.left = rect.left + (rect.width / 2) - (tooltip.offsetWidth / 2) + 'px';
            tooltip.style.top = rect.top - tooltip.offsetHeight - 10 + 'px';
            
            setTimeout(() => {
                tooltip.style.opacity = '1';
            }, 10);
            
            e.target.addEventListener('mouseleave', () => {
                tooltip.style.opacity = '0';
                setTimeout(() => {
                    tooltip.remove();
                }, 300);
            }, { once: true });
        });
    });
}

// リンクデバッグ用の初期化
function initLinkDebug() {
    // 全てのリンクにクリックイベントリスナーを追加
    const allLinks = document.querySelectorAll('a[href]');
    console.log(`Found ${allLinks.length} links on the page`);
    
    allLinks.forEach((link, index) => {
        console.log(`Link ${index + 1}:`, link.href, link.textContent);
        
        link.addEventListener('click', (e) => {
            console.log(`Link clicked:`, link.href, link.textContent);
            
            // 外部リンクやハッシュリンクでない場合は、リンクの動作を確認
            if (!link.href.startsWith('http') && !link.href.startsWith('#')) {
                console.log(`Internal link clicked:`, link.href);
                
                // リンク先のファイルが存在するかを確認
                fetch(link.href)
                    .then(response => {
                        if (response.ok) {
                            console.log(`Link target exists:`, link.href);
                        } else {
                            console.error(`Link target not found:`, link.href, response.status);
                        }
                    })
                    .catch(error => {
                        console.error(`Error checking link target:`, link.href, error);
                    });
            }
        });
    });
}

// ========================================
// ユーティリティ関数
// ========================================

// デバウンス関数（検索パフォーマンス向上用）
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// 検索入力にデバウンスを適用
if (searchInput) {
    const debouncedSearch = debounce((e) => {
        const searchTerm = e.target.value.toLowerCase().trim();
        if (searchTerm === '') {
            showAllCards();
        } else {
            filterCards(searchTerm);
            highlightSearchResults(searchTerm);
        }
    }, 300);
    
    searchInput.addEventListener('input', debouncedSearch);
}

// ========================================
// エクスポート（必要に応じて）
// ========================================
window.PowerShellDocs = {
    toggleTheme: () => {
        body.classList.toggle('dark-mode');
        updateThemeButton();
    },
    search: (term) => {
        searchInput.value = term;
        filterCards(term.toLowerCase());
    },
    showAll: showAllCards
};