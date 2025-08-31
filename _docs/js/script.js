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

// ローカルLLM設定
const localLLMConfig = {
    enabled: false,
    type: 'ollama', // 'ollama', 'lmstudio', 'llamacpp'
    endpoint: 'http://localhost:11434', // Ollamaのデフォルトエンドポイント
    model: 'gpt-oss:20b', // 使用するモデル（インストール済み）
    timeout: 30000, // タイムアウト（ミリ秒）
    fallbackToLocal: true // ローカルLLMが失敗した場合、ローカル回答にフォールバック
};

// 設定パネルの要素を取得する関数
function getSettingsElements() {
    const elements = {
        openChatbotSettings: document.getElementById('openChatbotSettings'),
        chatbotSettingsPanel: document.getElementById('chatbotSettingsPanel'),
        closeSettings: document.getElementById('closeSettings'),
        enableLocalLLM: document.getElementById('enableLocalLLM'),
        llmType: document.getElementById('llmType'),
        llmEndpoint: document.getElementById('llmEndpoint'),
        llmModel: document.getElementById('llmModel'),
        fallbackToLocal: document.getElementById('fallbackToLocal'),
        testConnection: document.getElementById('testConnection'),
        saveSettings: document.getElementById('saveSettings'),
        connectionStatus: document.getElementById('connectionStatus'),
        statusIndicator: document.getElementById('statusIndicator'),
        statusText: document.getElementById('statusText')
    };
    
    // デバッグ用：要素の存在確認
    console.log('設定関連要素の確認:');
    Object.entries(elements).forEach(([name, element]) => {
        console.log(`${name}:`, element);
    });
    
    return elements;
}

// 設定パネルの要素
let settingsElements = {};

// 設定パネルの表示/非表示
function setupSettingsEventListeners() {
    settingsElements = getSettingsElements();
    
    if (settingsElements.openChatbotSettings) {
        console.log('設定ボタンのイベントリスナーを設定中...');
        settingsElements.openChatbotSettings.addEventListener('click', () => {
            console.log('設定ボタンがクリックされました');
            if (settingsElements.chatbotSettingsPanel) {
                settingsElements.chatbotSettingsPanel.classList.add('show');
                loadSettingsToForm();
                console.log('設定パネルを表示しました');
            } else {
                console.error('設定パネルが見つかりません');
            }
        });
    } else {
        console.error('設定ボタンが見つかりません');
    }

    if (settingsElements.closeSettings) {
        settingsElements.closeSettings.addEventListener('click', () => {
            console.log('設定パネルを閉じます');
            if (settingsElements.chatbotSettingsPanel) {
                settingsElements.chatbotSettingsPanel.classList.remove('show');
            }
        });
    } else {
        console.error('設定パネル閉じるボタンが見つかりません');
    }
}

// 設定をフォームに読み込み
function loadSettingsToForm() {
    if (settingsElements.enableLocalLLM) {
        settingsElements.enableLocalLLM.checked = localLLMConfig.enabled;
        settingsElements.llmType.value = localLLMConfig.type;
        settingsElements.llmEndpoint.value = localLLMConfig.endpoint;
        settingsElements.llmModel.value = localLLMConfig.model;
        settingsElements.fallbackToLocal.checked = localLLMConfig.fallbackToLocal;
        
        updateConnectionStatus();
    } else {
        console.error('設定フォームの要素が見つかりません');
    }
}

// 設定を保存
function setupSaveSettingsListener() {
    if (settingsElements.saveSettings) {
        settingsElements.saveSettings.addEventListener('click', () => {
            localLLMConfig.enabled = settingsElements.enableLocalLLM.checked;
            localLLMConfig.type = settingsElements.llmType.value;
            localLLMConfig.endpoint = settingsElements.llmEndpoint.value;
            localLLMConfig.model = settingsElements.llmModel.value;
            localLLMConfig.fallbackToLocal = settingsElements.fallbackToLocal.checked;
            
            // ローカルストレージに保存
            localStorage.setItem('localLLMConfig', JSON.stringify(localLLMConfig));
            
            console.log('設定を保存しました:', localLLMConfig);
            
            // 設定パネルを閉じる
            if (settingsElements.chatbotSettingsPanel) {
                settingsElements.chatbotSettingsPanel.classList.remove('show');
                console.log('設定パネルを閉じました');
            }
            
            // 接続状態を更新
            updateConnectionStatus();
            
            // 成功メッセージを表示
            showNotification('設定を保存しました！', 'success');
        });
    } else {
        console.error('設定保存ボタンが見つかりません');
    }
}

// 接続テスト
function setupTestConnectionListener() {
    if (settingsElements.testConnection) {
        settingsElements.testConnection.addEventListener('click', async () => {
            setConnectionStatus('connecting', '接続テスト中...');
            
            try {
                const isConnected = await testLocalLLMConnection();
                if (isConnected) {
                    setConnectionStatus('connected', '接続成功');
                    showNotification('ローカルLLMへの接続が成功しました！', 'success');
                } else {
                    setConnectionStatus('disconnected', '接続失敗');
                    showNotification('ローカルLLMへの接続に失敗しました。設定を確認してください。', 'error');
                }
            } catch (error) {
                setConnectionStatus('disconnected', '接続エラー');
                showNotification('接続テスト中にエラーが発生しました。', 'error');
            }
        });
    } else {
        console.error('接続テストボタンが見つかりません');
    }
}

// 接続状態を設定
function setConnectionStatus(status, text) {
    if (settingsElements.statusIndicator && settingsElements.statusText) {
        settingsElements.statusIndicator.className = `status-indicator ${status}`;
        settingsElements.statusText.textContent = text;
    }
}

// 接続状態を更新
function updateConnectionStatus() {
    if (settingsElements.statusIndicator && settingsElements.statusText) {
        if (localLLMConfig.enabled) {
            setConnectionStatus('disconnected', '未接続');
        } else {
            setConnectionStatus('disconnected', '無効');
        }
    }
}

// 通知を表示
function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    
    // 通知のスタイルを設定
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 15px 20px;
        border-radius: 8px;
        color: white;
        font-weight: 600;
        z-index: 3000;
        animation: slideInRight 0.3s ease;
        max-width: 300px;
    `;
    
    // タイプ別の背景色
    switch (type) {
        case 'success':
            notification.style.background = 'var(--success-color)';
            break;
        case 'error':
            notification.style.background = 'var(--error-color)';
            break;
        default:
            notification.style.background = 'var(--primary-color)';
    }
    
    document.body.appendChild(notification);
    
    // 3秒後に自動削除
    setTimeout(() => {
        notification.style.animation = 'slideOutRight 0.3s ease';
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

// ローカルストレージから設定を読み込み
function loadSettingsFromStorage() {
    const savedConfig = localStorage.getItem('localLLMConfig');
    console.log('ローカルストレージから設定を読み込み中...');
    console.log('保存された設定:', savedConfig);
    
    if (savedConfig) {
        try {
            const config = JSON.parse(savedConfig);
            Object.assign(localLLMConfig, config);
            console.log('保存された設定を読み込みました:', localLLMConfig);
        } catch (error) {
            console.error('設定の読み込みに失敗しました:', error);
        }
    } else {
        console.log('保存された設定が見つかりません。デフォルト設定を使用します。');
        console.log('デフォルト設定:', localLLMConfig);
    }
}

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

// ローカルLLMとの通信
async function callLocalLLM(userMessage, context) {
    try {
        console.log('callLocalLLM呼び出し:', { userMessage, context, config: localLLMConfig });
        
        const prompt = generatePrompt(userMessage, context);
        console.log('生成されたプロンプト:', prompt);
        
        let result;
        switch (localLLMConfig.type) {
            case 'ollama':
                console.log('Ollamaを呼び出し中...');
                result = await callOllama(prompt);
                break;
            case 'lmstudio':
                console.log('LM Studioを呼び出し中...');
                result = await callLMStudio(prompt);
                break;
            case 'llamacpp':
                console.log('llama.cppを呼び出し中...');
                result = await callLlamaCpp(prompt);
                break;
            default:
                throw new Error('サポートされていないLLMタイプです');
        }
        
        console.log('LLM応答結果:', result);
        return result;
    } catch (error) {
        console.error('ローカルLLM呼び出しエラー:', error);
        return null;
    }
}

// Ollama API呼び出し
async function callOllama(prompt) {
    try {
        console.log('Ollama API呼び出し開始:', {
            endpoint: localLLMConfig.endpoint,
            model: localLLMConfig.model,
            prompt: prompt
        });
        
        const requestBody = {
            model: localLLMConfig.model,
            prompt: prompt,
            stream: false,
            options: {
                temperature: 0.7,
                top_p: 0.9,
                max_tokens: 1000
            }
        };
        
        console.log('Ollama API リクエスト:', requestBody);
        
        // CORSエラーを回避するため、XMLHttpRequestを使用
        return new Promise((resolve, reject) => {
            const xhr = new XMLHttpRequest();
            const timeoutId = setTimeout(() => {
                xhr.abort();
                reject(new Error('Ollama API呼び出しがタイムアウトしました。'));
            }, localLLMConfig.timeout || 30000);
            
            xhr.onload = function() {
                clearTimeout(timeoutId);
                if (xhr.status === 200) {
                    try {
                        const data = JSON.parse(xhr.responseText);
                        console.log('Ollama API レスポンスデータ:', data);
                        resolve(data.response);
                    } catch (e) {
                        reject(new Error('レスポンスの解析に失敗しました。'));
                    }
                } else {
                    reject(new Error(`Ollama API エラー: ${xhr.status} ${xhr.statusText}`));
                }
            };
            
            xhr.onerror = function() {
                clearTimeout(timeoutId);
                reject(new Error('ネットワークエラーが発生しました。'));
            };
            
            xhr.ontimeout = function() {
                clearTimeout(timeoutId);
                reject(new Error('リクエストがタイムアウトしました。'));
            };
            
            xhr.open('POST', `${localLLMConfig.endpoint}/api/generate`, true);
            xhr.setRequestHeader('Content-Type', 'application/json');
            xhr.send(JSON.stringify(requestBody));
        });
        
    } catch (error) {
        console.error('Ollama API呼び出しエラー:', error);
        
        // エラーの種類に応じて詳細なメッセージを生成
        let errorMessage = 'Ollama API呼び出しエラー';
        if (error.message.includes('タイムアウト')) {
            errorMessage = 'Ollama API呼び出しがタイムアウトしました。モデルの応答に時間がかかっている可能性があります。';
        } else if (error.message.includes('ネットワークエラー')) {
            errorMessage = 'Ollamaサービスに接続できません。サービスが起動しているか確認してください。';
        }
        
        const enhancedError = new Error(errorMessage);
        enhancedError.originalError = error;
        throw enhancedError;
    }
}

// LM Studio API呼び出し
async function callLMStudio(prompt) {
    const response = await fetch('http://localhost:1234/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            messages: [
                {
                    role: 'system',
                    content: 'あなたはPowerShell Driver Classesの専門家です。ユーザーの質問に日本語で丁寧に回答してください。'
                },
                {
                    role: 'user',
                    content: prompt
                }
            ],
            temperature: 0.7,
            max_tokens: 1000,
            stream: false
        })
    });

    if (!response.ok) {
        throw new Error(`LM Studio API エラー: ${response.status}`);
    }

    const data = await response.json();
    return data.choices[0].message.content;
}

// llama.cpp API呼び出し
async function callLlamaCpp(prompt) {
    const response = await fetch('http://localhost:8080/completion', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            prompt: prompt,
            n_predict: 1000,
            temperature: 0.7,
            top_p: 0.9,
            stop: ['</s>', 'Human:', 'Assistant:']
        })
    });

    if (!response.ok) {
        throw new Error(`llama.cpp API エラー: ${response.status}`);
    }

    const data = await response.json();
    return data.content;
}

// プロンプト生成
function generatePrompt(userMessage, context) {
    return `あなたはPowerShell Driver Classesの専門家です。

利用可能なクラス情報:
${Object.entries(ps1FileDatabase).map(([name, info]) => 
    `${name}: ${info.description} - 主なメソッド: ${info.methods.join(', ')}`
).join('\n')}

ユーザーの質問: ${userMessage}

コンテキスト: ${context}

上記の情報を基に、ユーザーの質問に日本語で丁寧に回答してください。具体的なコード例や使用法も含めて説明してください。`;
}

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
async function sendUserMessage() {
    const message = chatbotInput.value.trim();
    if (!message) return;
    
    // ユーザーメッセージを表示
    addMessage(message, 'user');
    chatbotInput.value = '';
    
    // 入力フィールドを無効化
    chatbotInput.disabled = true;
    sendMessage.disabled = true;
    
    // タイピングインジケーターを表示
    addTypingIndicator();
    
    // デバッグ情報を表示
    console.log('ローカルLLM設定:', localLLMConfig);
    console.log('質問内容:', message);
    
    try {
        let response;
        
        // ローカルLLMが有効で、設定されている場合
        if (localLLMConfig.enabled) {
            console.log('ローカルLLMを呼び出し中...');
            const context = `PowerShell Driver Classesの使い方について質問されています。`;
            response = await callLocalLLM(message, context);
            console.log('ローカルLLM応答:', response);
        } else {
            console.log('ローカルLLMが無効です。設定を確認してください。');
        }
        
        // ローカルLLMが失敗した場合、または無効な場合はローカル回答を使用
        if (!response && localLLMConfig.fallbackToLocal) {
            console.log('ローカル回答を使用します。');
            response = generateBotResponse(message);
        }
        
        // タイピングインジケーターを削除
        removeTypingIndicator();
        
        // 応答を表示
        addMessage(response || '申し訳ございません。回答を生成できませんでした。', 'bot');
        
    } catch (error) {
        console.error('エラー:', error);
        removeTypingIndicator();
        addMessage('エラーが発生しました。ローカル回答を使用します。', 'bot');
        
        // フォールバックとしてローカル回答を使用
        setTimeout(() => {
            const localResponse = generateBotResponse(message);
            addMessage(localResponse, 'bot');
        }, 500);
    } finally {
        // 入力フィールドを再有効化
        chatbotInput.disabled = false;
        sendMessage.disabled = false;
        chatbotInput.focus();
    }
}

// タイピングインジケーターを追加
function addTypingIndicator() {
    const typingDiv = document.createElement('div');
    typingDiv.className = 'message bot-message typing-indicator';
    typingDiv.id = 'typingIndicator';
    
    const typingContent = document.createElement('div');
    typingContent.className = 'message-content';
    typingContent.innerHTML = '<p>🤖 考え中...</p>';
    
    typingDiv.appendChild(typingContent);
    chatbotMessages.appendChild(typingDiv);
    chatbotMessages.scrollTop = chatbotMessages.scrollHeight;
}

// タイピングインジケーターを削除
function removeTypingIndicator() {
    const typingIndicator = document.getElementById('typingIndicator');
    if (typingIndicator) {
        typingIndicator.remove();
    }
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

// ボットの応答を生成（ローカルフォールバック用）
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

// ローカルLLM設定を更新
function updateLocalLLMConfig(newConfig) {
    Object.assign(localLLMConfig, newConfig);
    console.log('ローカルLLM設定を更新しました:', localLLMConfig);
}

// ローカルLLM接続テスト
async function testLocalLLMConnection() {
    try {
        console.log('ローカルLLM接続テスト開始...');
        
        // まず基本的なAPI接続をテスト
        const testResponse = await new Promise((resolve, reject) => {
            const xhr = new XMLHttpRequest();
            xhr.onload = function() {
                if (xhr.status === 200) {
                    resolve(xhr);
                } else {
                    reject(new Error(`API接続テスト失敗: ${xhr.status} ${xhr.statusText}`));
                }
            };
            xhr.onerror = function() {
                reject(new Error('API接続テストでネットワークエラーが発生しました。'));
            };
            xhr.open('GET', `${localLLMConfig.endpoint}/api/tags`, true);
            xhr.send();
        });
        
        console.log('API接続テスト成功');
        
        // 次に実際のLLM呼び出しをテスト
        const response = await callLocalLLM('テスト接続', '接続テスト');
        console.log('ローカルLLM接続成功:', response);
        return true;
    } catch (error) {
        console.error('ローカルLLM接続失敗:', error);
        return false;
    }
}

// ページ読み込み完了時の初期化
document.addEventListener('DOMContentLoaded', async () => {
    console.log('ページ初期化開始...');
    
    // 既存の初期化処理はそのまま
    initLinkDebug();
    
    // 保存された設定を読み込み
    loadSettingsFromStorage();
    
    // 設定パネルのイベントリスナーを設定
    setupSettingsEventListeners();
    setupSaveSettingsListener();
    setupTestConnectionListener();
    
    // チャットボットの初期化
    if (floatingActionButton && chatbotModal) {
        console.log('チャットボットが初期化されました');
        
        // ローカルLLM接続テスト（オプション）
        if (localLLMConfig.enabled) {
            console.log('ローカルLLMが有効です。接続テストを実行中...');
            const isConnected = await testLocalLLMConnection();
            if (isConnected) {
                console.log('ローカルLLMが利用可能です');
                addMessage('🤖 ローカルLLMが利用可能です。より詳細な回答が可能です。', 'bot');
            } else {
                console.log('ローカルLLMが利用できません。ローカル回答を使用します。');
                addMessage('⚠️ ローカルLLMが利用できません。ローカル回答を使用します。', 'bot');
            }
        } else {
            console.log('ローカルLLMが無効です。設定で有効化してください。');
            addMessage('ℹ️ ローカルLLMを使用するには、設定で有効化してください。', 'bot');
        }
    }
    
    // グローバル関数として公開（デバッグ用）
    window.PowerShellDocs.updateLocalLLMConfig = updateLocalLLMConfig;
    window.PowerShellDocs.testLocalLLMConnection = testLocalLLMConnection;
    window.PowerShellDocs.localLLMConfig = localLLMConfig;
    
    console.log('初期化完了。現在の設定:', localLLMConfig);
});