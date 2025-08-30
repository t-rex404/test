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

// ========================================
// チャットボット機能
// ========================================
const floatingActionButton = document.getElementById('floatingActionButton');
const chatbotModal = document.getElementById('chatbotModal');
const closeChatbot = document.getElementById('closeChatbot');
const chatbotInput = document.getElementById('chatbotInput');
const sendMessage = document.getElementById('sendMessage');
const chatbotMessages = document.getElementById('chatbotMessages');

// PS1ファイルの情報データベース
const ps1FileDatabase = {
    'WebDriver': {
        description: 'ブラウザ操作の基底クラス',
        methods: ['Navigate', 'FindElement', 'Click', 'SendKeys', 'Screenshot', 'GetCookies'],
        usage: 'ChromeDriverやEdgeDriverの親クラスとして使用',
        example: 'WebDriverを継承してカスタムドライバーを作成できます'
    },
    'ChromeDriver': {
        description: 'Google Chromeブラウザを自動操作',
        methods: ['StartChrome', 'SetWindowSize', 'ExecuteScript', 'WaitForElement'],
        usage: 'Chromeブラウザの自動化に使用',
        example: 'ChromeDriverをインスタンス化してブラウザを起動し、Webサイトを操作できます'
    },
    'EdgeDriver': {
        description: 'Microsoft Edgeブラウザを自動操作',
        methods: ['StartEdge', 'SetWindowSize', 'ExecuteScript', 'WaitForElement'],
        usage: 'Edgeブラウザの自動化に使用',
        example: 'EdgeDriverをインスタンス化してブラウザを起動し、Webサイトを操作できます'
    },
    'WordDriver': {
        description: 'Microsoft Wordを自動操作',
        methods: ['CreateDocument', 'AddText', 'SetFont', 'InsertTable', 'SaveDocument'],
        usage: 'Word文書の自動作成・編集に使用',
        example: 'WordDriverで文書を作成し、テキストや表を挿入して保存できます'
    },
    'ExcelDriver': {
        description: 'Microsoft Excelを自動操作',
        methods: ['OpenWorkbook', 'SetCellValue', 'FormatCell', 'CreateChart', 'SaveWorkbook'],
        usage: 'Excelファイルの自動作成・編集に使用',
        example: 'ExcelDriverでワークブックを開き、セルに値を設定して保存できます'
    },
    'PowerPointDriver': {
        description: 'Microsoft PowerPointを自動操作',
        methods: ['CreatePresentation', 'AddSlide', 'InsertShape', 'SetText', 'SavePresentation'],
        usage: 'PowerPointプレゼンテーションの自動作成に使用',
        example: 'PowerPointDriverでプレゼンテーションを作成し、スライドや図形を追加できます'
    },
    'OracleDriver': {
        description: 'Oracleデータベースを操作',
        methods: ['Connect', 'ExecuteQuery', 'ExecuteNonQuery', 'BeginTransaction', 'Commit'],
        usage: 'Oracleデータベースへの接続・操作に使用',
        example: 'OracleDriverでデータベースに接続し、SQLクエリを実行できます'
    },
    'Common': {
        description: '共通機能を提供するユーティリティクラス',
        methods: ['WriteLog', 'HandleError', 'GetErrorCode', 'FormatMessage'],
        usage: '全ドライバークラスで使用する共通機能',
        example: 'Commonクラスのログ出力やエラー処理機能を活用できます'
    }
};

// フローティングアクションボタンのクリックイベント
floatingActionButton?.addEventListener('click', () => {
    chatbotModal.classList.add('show');
    chatbotInput.focus();
});

// チャットボットを閉じる
closeChatbot?.addEventListener('click', () => {
    chatbotModal.classList.remove('show');
});

// モーダル外クリックで閉じる
chatbotModal?.addEventListener('click', (e) => {
    if (e.target === chatbotModal) {
        chatbotModal.classList.remove('show');
    }
});

// Enterキーでメッセージ送信
chatbotInput?.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        sendUserMessage();
    }
});

// 送信ボタンのクリックイベント
sendMessage?.addEventListener('click', sendUserMessage);

// ユーザーメッセージを送信
function sendUserMessage() {
    const message = chatbotInput.value.trim();
    if (!message) return;
    
    // ユーザーメッセージを表示
    addMessage(message, 'user');
    chatbotInput.value = '';
    
    // ボットの応答を生成
    setTimeout(() => {
        const response = generateBotResponse(message);
        addMessage(response, 'bot');
    }, 500);
}

// メッセージをチャットに追加
function addMessage(content, sender) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${sender}-message`;
    
    const messageContent = document.createElement('div');
    messageContent.className = 'message-content';
    
    if (typeof content === 'string') {
        messageContent.innerHTML = `<p>${content}</p>`;
    } else {
        messageContent.innerHTML = content;
    }
    
    messageDiv.appendChild(messageContent);
    chatbotMessages.appendChild(messageDiv);
    
    // スクロールを最下部に
    chatbotMessages.scrollTop = chatbotMessages.scrollHeight;
}

// ボットの応答を生成
function generateBotResponse(userMessage) {
    const message = userMessage.toLowerCase();
    
    // 特定のキーワードに対する応答
    if (message.includes('こんにちは') || message.includes('hello')) {
        return 'こんにちは！PowerShell Driver Classesについて何でもお聞きください。';
    }
    
    if (message.includes('使い方') || message.includes('how to use')) {
        return 'どのクラスの使い方を知りたいですか？例えば「ChromeDriverの使い方を教えて」のように質問してください。';
    }
    
    if (message.includes('メソッド') || message.includes('method')) {
        return 'どのクラスのメソッドについて知りたいですか？具体的なクラス名を教えてください。';
    }
    
    // 各ドライバークラスに関する質問
    for (const [className, info] of Object.entries(ps1FileDatabase)) {
        if (message.includes(className.toLowerCase()) || message.includes(className.replace('Driver', '').toLowerCase())) {
            return generateClassInfo(className, info);
        }
    }
    
    // 一般的な質問に対する応答
    if (message.includes('エラー') || message.includes('error')) {
        return 'エラーが発生した場合は、CommonクラスのWriteLogメソッドでログを確認し、GetErrorCodeでエラーコードを取得してください。';
    }
    
    if (message.includes('ログ') || message.includes('log')) {
        return 'ログ出力にはCommonクラスのWriteLogメソッドを使用します。詳細なログでデバッグを効率化できます。';
    }
    
    if (message.includes('インストール') || message.includes('install')) {
        return 'PowerShell 5.1以上が必要です。各ドライバークラスを使用するには、対応するアプリケーション（Chrome、Office等）のインストールが必要です。';
    }
    
    // デフォルト応答
    return '申し訳ございません。もう少し具体的に質問していただけますか？例えば「ChromeDriverの使い方」「WordDriverで文書を作成する方法」など。';
}

// クラス情報を生成
function generateClassInfo(className, info) {
    return `
        <h4>${className}について</h4>
        <p><strong>説明:</strong> ${info.description}</p>
        <p><strong>主なメソッド:</strong></p>
        <ul>
            ${info.methods.map(method => `<li>${method}</li>`).join('')}
        </ul>
        <p><strong>使用例:</strong> ${info.example}</p>
        <p><strong>詳細:</strong> <a href="pages/${className.toLowerCase()}.html" target="_blank">${className}の詳細ページ</a>をご確認ください。</p>
    `;
}

// ページ読み込み完了時の初期化
document.addEventListener('DOMContentLoaded', () => {
    // 既存の初期化処理はそのまま
    initLinkDebug();
    
    // チャットボットの初期化
    if (floatingActionButton && chatbotModal) {
        console.log('チャットボットが初期化されました');
    }
});