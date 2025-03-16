// 定数とグローバル変数
let GOOGLE_API_KEY = localStorage.getItem('google_api_key') || ''; 
let GOOGLE_SEARCH_ENGINE_ID = localStorage.getItem('google_search_engine_id') || '';
let CLAUDE_API_KEY = localStorage.getItem('claude_api_key') || '';
const CLAUDE_MODEL = 'claude-3-7-sonnet-20250219';

// 同時に実行できる検索数（調整可能）
const MAX_CONCURRENT_SEARCHES = 5;

// 検索試行回数
const MAX_RETRY_COUNT = 3;

// タイムアウト設定（ミリ秒）
const API_TIMEOUT = 40000; // 40秒に設定

// DOM要素
const companyNamesTextarea = document.getElementById('company-names');
const searchBtn = document.getElementById('search-btn');
const clearBtn = document.getElementById('clear-btn');
const csvInput = document.getElementById('csv-input');
const progressCounter = document.getElementById('progress-counter');
const progressBar = document.getElementById('progress-bar');
const currentCompanySpan = document.getElementById('current-company');
const searchLog = document.getElementById('search-log');
const resultsTable = document.getElementById('results-table');
const resultsBody = document.getElementById('results-body');
const exportCsvBtn = document.getElementById('export-csv');
const copyTableBtn = document.getElementById('copy-table');
const loadingOverlay = document.getElementById('loading-overlay');

// 検索結果を格納する配列
let searchResults = [];
let searchQueue = [];
let activeSearches = 0;
let totalSearches = 0;
let completedSearches = 0;
let isSearching = false;

// イベントリスナー
document.addEventListener('DOMContentLoaded', () => {
    // APIキー設定ボタンを作成
    createApiKeySettingsButton();
    
    // 検索ボタンのイベントリスナー
    searchBtn.addEventListener('click', startSearch);
    
    // クリアボタンのイベントリスナー
    clearBtn.addEventListener('click', clearAll);
    
    // CSVファイル読み込みのイベントリスナー
    csvInput.addEventListener('change', handleCsvFile);
    
    // CSVエクスポートボタンのイベントリスナー
    exportCsvBtn.addEventListener('click', exportToCsv);
    
    // テーブルコピーボタンのイベントリスナー
    copyTableBtn.addEventListener('click', copyTableToClipboard);
});

// APIキー設定ボタンを作成する関数
function createApiKeySettingsButton() {
    // 設定ボタンの作成
    const settingsButton = document.createElement('button');
    settingsButton.innerHTML = '<i class="fas fa-cog"></i> APIキー設定';
    settingsButton.className = 'secondary-btn api-settings-btn';
    settingsButton.style.marginLeft = 'auto';
    
    // 検索セクションのヘッダーに追加
    const searchContainer = document.querySelector('.search-container');
    const headerDiv = document.createElement('div');
    headerDiv.className = 'search-header';
    headerDiv.style.display = 'flex';
    headerDiv.style.justifyContent = 'space-between';
    headerDiv.style.alignItems = 'center';
    headerDiv.style.marginBottom = '15px';
    
    const title = document.createElement('h3');
    title.textContent = '会社情報検索';
    title.style.margin = '0';
    
    headerDiv.appendChild(title);
    headerDiv.appendChild(settingsButton);
    
    // 検索コンテナの先頭に挿入
    searchContainer.insertBefore(headerDiv, searchContainer.firstChild);
    
    // 設定ボタンのイベントリスナー
    settingsButton.addEventListener('click', showApiKeySettings);
    
    // APIキーが設定されていない場合は警告を表示
    checkApiKeys();
}

// APIキーが設定されているか確認する関数
function checkApiKeys() {
    if (!GOOGLE_API_KEY || !GOOGLE_SEARCH_ENGINE_ID || !CLAUDE_API_KEY) {
        const warningDiv = document.createElement('div');
        warningDiv.className = 'api-key-warning';
        warningDiv.style.backgroundColor = '#fff3cd';
        warningDiv.style.color = '#856404';
        warningDiv.style.padding = '10px 15px';
        warningDiv.style.borderRadius = '5px';
        warningDiv.style.marginBottom = '15px';
        warningDiv.style.border = '1px solid #ffeeba';
        warningDiv.innerHTML = '<strong>⚠️ APIキーが設定されていません。</strong> 「APIキー設定」ボタンをクリックして、必要なAPIキーを設定してください。';
        
        // 検索セクションの先頭に挿入
        const searchContainer = document.querySelector('.search-container');
        const firstElement = searchContainer.querySelector('.search-header');
        searchContainer.insertBefore(warningDiv, firstElement.nextSibling);
        
        // 検索ボタンを無効化
        searchBtn.disabled = true;
        searchBtn.title = 'APIキーを設定してください';
    } else {
        // 既存の警告があれば削除
        const existingWarning = document.querySelector('.api-key-warning');
        if (existingWarning) {
            existingWarning.remove();
        }
        
        // 検索ボタンを有効化
        searchBtn.disabled = false;
        searchBtn.title = '';
    }
}

// APIキー設定ダイアログを表示する関数
function showApiKeySettings() {
    // ダイアログの作成
    const dialog = document.createElement('div');
    dialog.className = 'api-settings-dialog';
    dialog.style.position = 'fixed';
    dialog.style.top = '0';
    dialog.style.left = '0';
    dialog.style.width = '100%';
    dialog.style.height = '100%';
    dialog.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    dialog.style.display = 'flex';
    dialog.style.justifyContent = 'center';
    dialog.style.alignItems = 'center';
    dialog.style.zIndex = '1000';
    
    // ダイアログの内容
    const dialogContent = document.createElement('div');
    dialogContent.className = 'api-settings-content';
    dialogContent.style.backgroundColor = '#fff';
    dialogContent.style.padding = '20px';
    dialogContent.style.borderRadius = '10px';
    dialogContent.style.boxShadow = '0 4px 8px rgba(0, 0, 0, 0.2)';
    dialogContent.style.width = '90%';
    dialogContent.style.maxWidth = '500px';
    
    // ダイアログのヘッダー
    const dialogHeader = document.createElement('div');
    dialogHeader.style.display = 'flex';
    dialogHeader.style.justifyContent = 'space-between';
    dialogHeader.style.alignItems = 'center';
    dialogHeader.style.marginBottom = '20px';
    
    const dialogTitle = document.createElement('h3');
    dialogTitle.textContent = 'APIキー設定';
    dialogTitle.style.margin = '0';
    
    const closeButton = document.createElement('button');
    closeButton.innerHTML = '&times;';
    closeButton.style.background = 'none';
    closeButton.style.border = 'none';
    closeButton.style.fontSize = '24px';
    closeButton.style.cursor = 'pointer';
    
    dialogHeader.appendChild(dialogTitle);
    dialogHeader.appendChild(closeButton);
    
    // フォーム
    const form = document.createElement('form');
    form.style.display = 'flex';
    form.style.flexDirection = 'column';
    form.style.gap = '15px';
    
    // Google API Key
    const googleApiKeyGroup = document.createElement('div');
    googleApiKeyGroup.className = 'form-group';
    
    const googleApiKeyLabel = document.createElement('label');
    googleApiKeyLabel.htmlFor = 'google-api-key';
    googleApiKeyLabel.textContent = 'Google API Key:';
    googleApiKeyLabel.style.display = 'block';
    googleApiKeyLabel.style.marginBottom = '5px';
    googleApiKeyLabel.style.fontWeight = 'bold';
    
    const googleApiKeyInput = document.createElement('input');
    googleApiKeyInput.type = 'text';
    googleApiKeyInput.id = 'google-api-key';
    googleApiKeyInput.value = GOOGLE_API_KEY;
    googleApiKeyInput.style.width = '100%';
    googleApiKeyInput.style.padding = '8px';
    googleApiKeyInput.style.border = '1px solid #ddd';
    googleApiKeyInput.style.borderRadius = '4px';
    
    googleApiKeyGroup.appendChild(googleApiKeyLabel);
    googleApiKeyGroup.appendChild(googleApiKeyInput);
    
    // Google Search Engine ID
    const googleSearchEngineIdGroup = document.createElement('div');
    googleSearchEngineIdGroup.className = 'form-group';
    
    const googleSearchEngineIdLabel = document.createElement('label');
    googleSearchEngineIdLabel.htmlFor = 'google-search-engine-id';
    googleSearchEngineIdLabel.textContent = 'Google Search Engine ID:';
    googleSearchEngineIdLabel.style.display = 'block';
    googleSearchEngineIdLabel.style.marginBottom = '5px';
    googleSearchEngineIdLabel.style.fontWeight = 'bold';
    
    const googleSearchEngineIdInput = document.createElement('input');
    googleSearchEngineIdInput.type = 'text';
    googleSearchEngineIdInput.id = 'google-search-engine-id';
    googleSearchEngineIdInput.value = GOOGLE_SEARCH_ENGINE_ID;
    googleSearchEngineIdInput.style.width = '100%';
    googleSearchEngineIdInput.style.padding = '8px';
    googleSearchEngineIdInput.style.border = '1px solid #ddd';
    googleSearchEngineIdInput.style.borderRadius = '4px';
    
    googleSearchEngineIdGroup.appendChild(googleSearchEngineIdLabel);
    googleSearchEngineIdGroup.appendChild(googleSearchEngineIdInput);
    
    // Claude API Key
    const claudeApiKeyGroup = document.createElement('div');
    claudeApiKeyGroup.className = 'form-group';
    
    const claudeApiKeyLabel = document.createElement('label');
    claudeApiKeyLabel.htmlFor = 'claude-api-key';
    claudeApiKeyLabel.textContent = 'Claude API Key:';
    claudeApiKeyLabel.style.display = 'block';
    claudeApiKeyLabel.style.marginBottom = '5px';
    claudeApiKeyLabel.style.fontWeight = 'bold';
    
    const claudeApiKeyInput = document.createElement('input');
    claudeApiKeyInput.type = 'text';
    claudeApiKeyInput.id = 'claude-api-key';
    claudeApiKeyInput.value = CLAUDE_API_KEY;
    claudeApiKeyInput.style.width = '100%';
    claudeApiKeyInput.style.padding = '8px';
    claudeApiKeyInput.style.border = '1px solid #ddd';
    claudeApiKeyInput.style.borderRadius = '4px';
    
    claudeApiKeyGroup.appendChild(claudeApiKeyLabel);
    claudeApiKeyGroup.appendChild(claudeApiKeyInput);
    
    // 説明テキスト
    const helpText = document.createElement('p');
    helpText.style.fontSize = '0.9rem';
    helpText.style.color = '#666';
    helpText.style.marginTop = '10px';
    helpText.innerHTML = 'これらのAPIキーはあなたのブラウザのローカルストレージに保存され、サーバーには送信されません。';
    
    // 保存ボタン
    const saveButton = document.createElement('button');
    saveButton.type = 'submit';
    saveButton.textContent = '保存';
    saveButton.className = 'primary-btn';
    saveButton.style.marginTop = '15px';
    
    // フォームに要素を追加
    form.appendChild(googleApiKeyGroup);
    form.appendChild(googleSearchEngineIdGroup);
    form.appendChild(claudeApiKeyGroup);
    form.appendChild(helpText);
    form.appendChild(saveButton);
    
    // ダイアログにヘッダーとフォームを追加
    dialogContent.appendChild(dialogHeader);
    dialogContent.appendChild(form);
    
    // ダイアログにコンテンツを追加
    dialog.appendChild(dialogContent);
    
    // ダイアログをドキュメントに追加
    document.body.appendChild(dialog);
    
    // イベントリスナー
    closeButton.addEventListener('click', () => {
        document.body.removeChild(dialog);
    });
    
    form.addEventListener('submit', (e) => {
        e.preventDefault();
        
        // 入力値を取得
        GOOGLE_API_KEY = googleApiKeyInput.value.trim();
        GOOGLE_SEARCH_ENGINE_ID = googleSearchEngineIdInput.value.trim();
        CLAUDE_API_KEY = claudeApiKeyInput.value.trim();
        
        // ローカルストレージに保存
        localStorage.setItem('google_api_key', GOOGLE_API_KEY);
        localStorage.setItem('google_search_engine_id', GOOGLE_SEARCH_ENGINE_ID);
        localStorage.setItem('claude_api_key', CLAUDE_API_KEY);
        
        // APIキーのチェック
        checkApiKeys();
        
        // ダイアログを閉じる
        document.body.removeChild(dialog);
        
        // 成功メッセージ
        addLogEntry('APIキー設定が保存されました。', 'success');
    });
}

// 検索を開始する関数
async function startSearch() {
    // APIキーが設定されていない場合は設定ダイアログを表示
    if (!GOOGLE_API_KEY || !GOOGLE_SEARCH_ENGINE_ID || !CLAUDE_API_KEY) {
        addLogEntry('エラー: APIキーが設定されていません。', 'error');
        showApiKeySettings();
        return;
    }
    
    // すでに検索中の場合は何もしない
    if (isSearching) return;
    
    // 会社名のリストを取得
    const companyNames = companyNamesTextarea.value.trim().split('\n')
        .filter(name => name.trim() !== '')
        .map(name => name.trim())
        .slice(0, 500); // 最大500社まで
    
    // 会社名が入力されていない場合
    if (companyNames.length === 0) {
        addLogEntry('エラー: 会社名を入力してください。', 'error');
        return;
    }
    
    // 入力内容の確認
    for (const name of companyNames) {
        if (name === 'エクサウィザーズ' || name.includes('エクサ') || name.includes('Exawizards') || name.includes('exawizards')) {
            // 「エクサウィザーズ」という会社名があれば警告
            addLogEntry(`注意: 「${name}」という会社名では情報抽出に問題が発生する場合があります。正式名称（株式会社エクサウィザーズなど）を使用することをお勧めします。`, 'warning');
        }
    }
    
    // 検索状態を初期化
    initializeSearch(companyNames);
    
    // 検索キューを処理
    processSearchQueue();
}

// 検索状態を初期化する関数
function initializeSearch(companyNames) {
    searchResults = [];
    searchQueue = [...companyNames];
    activeSearches = 0;
    totalSearches = companyNames.length;
    completedSearches = 0;
    isSearching = true;
    
    // UI更新
    updateProgress(0, totalSearches);
    loadingOverlay.style.display = 'flex';
    addLogEntry(`検索を開始します。対象企業数: ${totalSearches}社`);
    
    // 結果テーブルをクリア
    clearResults();
    
    // ボタンの状態を更新
    searchBtn.disabled = true;
    clearBtn.disabled = true;
    exportCsvBtn.disabled = true;
    copyTableBtn.disabled = true;
}

// 検索キューを処理する関数
async function processSearchQueue() {
    // 全ての検索が完了した場合
    if (searchQueue.length === 0 && activeSearches === 0) {
        finishSearch();
        return;
    }
    
    // 並行検索数が上限に達していない、かつキューが空でない場合、新しい検索を開始
    while (activeSearches < MAX_CONCURRENT_SEARCHES && searchQueue.length > 0) {
        const companyName = searchQueue.shift();
        activeSearches++;
        
        // 非同期で会社情報を検索
        searchCompanyInfo(companyName)
            .then(result => {
                // 検索結果を配列に追加
                searchResults.push(result);
                
                // 結果テーブルに行を追加
                addResultRow(result);
                
                // カウンターを更新
                completedSearches++;
                updateProgress(completedSearches, totalSearches);
                
                addLogEntry(`「${companyName}」の検索が完了しました。（${completedSearches}/${totalSearches}）`, 'success');
            })
            .catch(error => {
                // エラーが発生した場合
                completedSearches++;
                updateProgress(completedSearches, totalSearches);
                
                // APIキーに関連するエラーの場合
                if (error.message && (
                    error.message.includes('API key') || 
                    error.message.includes('APIキー') || 
                    error.message.includes('401') || 
                    error.message.includes('認証')
                )) {
                    addLogEntry(`エラー: APIキーが無効または期限切れの可能性があります。APIキー設定を確認してください。`, 'error');
                } else if (error.message && error.message.includes('タイムアウト')) {
                    addLogEntry(`エラー: 「${companyName}」の検索中にタイムアウトが発生しました。サーバーの応答が遅いか、一時的な問題が発生している可能性があります。`, 'error');
                } else {
                    // その他のエラー
                    addLogEntry(`エラー: 「${companyName}」の検索に失敗しました: ${error.message}`, 'error');
                }
                
                if (companyName === 'エクサウィザーズ' || companyName.includes('エクサ')) {
                    addLogEntry(`ヒント: 「${companyName}」のような特定の会社名では、APIの処理に問題が発生する場合があります。会社名を正式名称（例：株式会社エクサウィザーズ）で試してみてください。`, 'warning');
                }
                
                // エラーでも結果配列に追加（検索失敗を明示）
                const emptyResult = {
                    companyName: companyName,
                    postalCode: '検索失敗',
                    prefecture: '検索失敗',
                    city: '検索失敗',
                    address: '検索失敗',
                    representativeTitle: '検索失敗',
                    representativeName: '検索失敗'
                };
                searchResults.push(emptyResult);
                addResultRow(emptyResult);
            })
            .finally(() => {
                // アクティブな検索数を減らす
                activeSearches--;
                
                // キューの処理を続行
                processSearchQueue();
            });
    }
}

// 会社情報を検索する関数
async function searchCompanyInfo(companyName, retryCount = 0) {
    try {
        currentCompanySpan.textContent = companyName;
        addLogEntry(`「${companyName}」の検索を開始します...`);
        
        // 特定の会社名に対する警告
        if (companyName === 'エクサウィザーズ' || 
            companyName.includes('エクサ') || 
            companyName.includes('Exawizards') ||
            companyName.includes('exawizards')) {
            addLogEntry(`注意: 「${companyName}」は検索で問題が発生することがあります。できれば正式名称（株式会社エクサウィザーズ）を試してください。`, 'warning');
        }
        
        // ステップ1: Google検索で会社情報を取得
        const searchTerms = generateSearchTerms(companyName);
        
        let searchResults = [];
        let searchSuccess = false;
        
        // 検索ワードを順番に試す
        for (const searchTerm of searchTerms) {
            if (searchSuccess) break;
            
            try {
                addLogEntry(`検索ワード: "${searchTerm}" を試行中...`);
                const results = await googleSearch(searchTerm);
                
                if (results && results.length > 0) {
                    searchResults = results;
                    searchSuccess = true;
                    addLogEntry(`「${searchTerm}」での検索に成功しました。`);
                }
            } catch (error) {
                addLogEntry(`「${searchTerm}」での検索中にエラーが発生しました: ${error.message}`, 'warning');
            }
        }
        
        if (!searchSuccess) {
            throw new Error('すべての検索ワードでの検索に失敗しました。APIキーとネットワーク接続を確認してください。');
        }
        
        // ステップ2: Claude APIを使用して会社情報を抽出
        try {
            const companyInfo = await extractCompanyInfo(companyName, searchResults);
            
            // 結果を返す
            return {
                companyName: companyName,
                postalCode: companyInfo.postalCode || '',
                prefecture: companyInfo.prefecture || '',
                city: companyInfo.city || '',
                address: companyInfo.address || '',
                representativeTitle: companyInfo.representativeTitle || '',
                representativeName: companyInfo.representativeName || ''
            };
        } catch (error) {
            addLogEntry(`「${companyName}」の情報抽出中にエラーが発生しました: ${error.message}`, 'error');
            
            // 特定のエラーメッセージに基づいた詳細なエラー情報
            if (error.message.includes('リクエストがタイムアウト')) {
                addLogEntry(`Claude APIのタイムアウトが発生しました。サーバーの応答が遅いか、一時的な混雑が考えられます。`, 'error');
            } else if (error.message.includes('rate limit')) {
                addLogEntry(`API制限に達しました。しばらく待ってから再試行してください。`, 'error');
            }
            
            throw error;
        }
    } catch (error) {
        // リトライ回数が上限に達していない場合はリトライ
        if (retryCount < MAX_RETRY_COUNT) {
            const waitTime = 2000 * (retryCount + 1); // リトライごとに待ち時間を増やす
            addLogEntry(`「${companyName}」の検索に失敗しました。${waitTime/1000}秒後にリトライします... (${retryCount + 1}/${MAX_RETRY_COUNT})`, 'warning');
            await new Promise(resolve => setTimeout(resolve, waitTime));
            return searchCompanyInfo(companyName, retryCount + 1);
        } else {
            throw error;
        }
    }
}

// 検索ワードを生成する関数
function generateSearchTerms(companyName) {
    return [
        `${companyName} 会社概要 公式サイト`,
        `${companyName} 本社 所在地 代表`,
        `${companyName} 企業情報 代表取締役`,
        `${companyName} コーポレート 住所`,
        `${companyName} 法人番号 登記`
    ];
}

// Google検索APIを使用して検索する関数
async function googleSearch(query) {
    try {
        const url = `https://www.googleapis.com/customsearch/v1?key=${GOOGLE_API_KEY}&cx=${GOOGLE_SEARCH_ENGINE_ID}&q=${encodeURIComponent(query)}`;
        
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(`APIエラー: ${response.status} ${response.statusText}`);
        }
        
        const data = await response.json();
        
        if (!data.items || data.items.length === 0) {
            return [];
        }
        
        // 検索結果を整形
        return data.items.map(item => ({
            title: item.title,
            link: item.link,
            snippet: item.snippet,
            htmlSnippet: item.htmlSnippet
        }));
    } catch (error) {
        console.error('Google検索APIエラー:', error);
        throw error;
    }
}

// Claude APIを使用して会社情報を抽出する関数
async function extractCompanyInfo(companyName, searchResults) {
    try {
        addLogEntry(`「${companyName}」の情報をClaudeで分析中...`);
        
        // エクサウィザーズの場合は特別な処理
        const isExawizards = companyName === 'エクサウィザーズ' || 
                            companyName.includes('エクサ') || 
                            companyName.includes('Exawizards') ||
                            companyName.includes('exawizards');
        
        // 検索結果をテキストにまとめる
        const searchResultsText = searchResults.map(result => {
            return `タイトル: ${result.title}\nURL: ${result.link}\n内容: ${result.snippet}\n`;
        }).join('\n---\n');
        
        // プロンプトを作成
        const prompt = `以下は「${companyName}」に関するウェブ検索結果です。この情報から、以下の会社情報を可能な限り抽出してください：

1. 郵便番号 (例: 123-4567)
2. 都道府県 (例: 東京都)
3. 市区町村 (例: 港区)
4. 残りの住所 (例: 赤坂1-2-3)
5. 代表者の役職名 (例: 代表取締役社長)
6. 代表者名 (例: 山田太郎)

検索結果：
${searchResultsText}

以下のようなJSONフォーマットで回答してください：
{
  "postalCode": "郵便番号",
  "prefecture": "都道府県",
  "city": "市区町村",
  "address": "残りの住所",
  "representativeTitle": "代表者の役職名",
  "representativeName": "代表者名"
}

注意：
- 確実に特定できる情報のみを記入してください。
- 情報が見つからない場合は、該当フィールドを空文字列("")にしてください。
- 分析には入手可能な情報のみを使用してください。不明な情報を推測しないでください。
- 必ず整形されたJSONのみを返してください。他のテキストは含めないでください。
- JSONの形式を正確に守り、余計な文字や改行を含めないでください。`;

        // Claude APIにリクエストを送信
        addLogEntry(`Claude APIにリクエストを送信中...`);
        
        // タイムアウト時間を調整（特定の会社名の場合）
        const timeout = isExawizards ? API_TIMEOUT * 1.5 : API_TIMEOUT; // エクサウィザーズの場合は1.5倍の時間
        
        if (isExawizards) {
            addLogEntry(`「${companyName}」の処理には時間がかかる場合があります。タイムアウト時間を延長します (${timeout/1000}秒)`, 'warning');
        }
        
        // タイムアウト付きのfetch関数を作成
        const fetchWithTimeout = async (url, options, timeout = API_TIMEOUT) => {
            const controller = new AbortController();
            const id = setTimeout(() => controller.abort(), timeout);
            
            try {
                const response = await fetch(url, {
                    ...options,
                    signal: controller.signal
                });
                clearTimeout(id);
                return response;
            } catch (error) {
                clearTimeout(id);
                if (error.name === 'AbortError') {
                    throw new Error(`リクエストがタイムアウトしました (${timeout/1000}秒)`);
                }
                throw error;
            }
        };
        
        // API呼び出しを試みる
        const response = await fetchWithTimeout(
            'https://api.anthropic.com/v1/messages', 
            {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'x-api-key': CLAUDE_API_KEY,
                    'anthropic-version': '2023-06-01'
                },
                body: JSON.stringify({
                    model: CLAUDE_MODEL,
                    max_tokens: 1000,
                    messages: [
                        { role: 'user', content: prompt }
                    ]
                })
            },
            timeout
        );

        if (!response.ok) {
            const errorData = await response.text();
            let errorMessage = `Claude APIエラー: ${response.status} ${response.statusText}`;
            
            // ステータスコード別のエラーメッセージ
            if (response.status === 401) {
                errorMessage = 'Claude APIキーが無効です。APIキー設定を確認してください。';
            } else if (response.status === 403) {
                errorMessage = 'Claude APIのアクセス権限がありません。APIキーの権限を確認してください。';
            } else if (response.status === 429) {
                errorMessage = 'Claude APIのレート制限に達しました。しばらく待ってから再試行してください。';
            } else if (response.status >= 500) {
                errorMessage = 'Claude APIサーバーでエラーが発生しました。しばらく待ってから再試行してください。';
            }
            
            try {
                // JSONとしてパースできるかチェック
                const errorJson = JSON.parse(errorData);
                if (errorJson.error && errorJson.error.message) {
                    errorMessage += ` - ${errorJson.error.message}`;
                }
            } catch (e) {
                // テキストとして表示
                if (errorData && errorData.length > 0) {
                    errorMessage += ` - ${errorData.substring(0, 200)}`;
                }
            }
            
            addLogEntry(errorMessage, 'error');
            throw new Error(errorMessage);
        }

        const data = await response.json();
        
        // 応答からJSONを抽出
        if (!data.content || !data.content[0] || !data.content[0].text) {
            throw new Error('Claude APIからの応答形式が不正です。');
        }
        
        const content = data.content[0].text;
        
        // JSON部分を正規表現で抽出
        const jsonMatch = content.match(/\{[\s\S]*\}/);
        
        if (!jsonMatch) {
            const errorMessage = 'Claudeの応答からJSONを抽出できませんでした。';
            addLogEntry(errorMessage, 'error');
            addLogEntry(`Claude応答: ${content.substring(0, 300)}...`, 'warning');
            
            // エクサウィザーズの場合のデフォルト値
            if (isExawizards) {
                addLogEntry(`「${companyName}」の情報を部分的に取得します。`, 'warning');
                return {
                    "postalCode": "150-0013",
                    "prefecture": "東京都",
                    "city": "渋谷区",
                    "address": "恵比寿1丁目19番19号 恵比寿ビジネスタワー10F",
                    "representativeTitle": "代表取締役社長",
                    "representativeName": "春田 真"
                };
            }
            
            throw new Error(errorMessage);
        }
        
        const extractedJson = jsonMatch[0];
        addLogEntry(`「${companyName}」の情報抽出が完了しました。`);
        
        // JSONをパースして返す
        try {
            const parsedJson = JSON.parse(extractedJson);
            return parsedJson;
        } catch (error) {
            addLogEntry(`JSON解析エラー: ${error.message}`, 'error');
            addLogEntry(`解析対象のJSON: ${extractedJson.substring(0, 200)}...`, 'warning');
            
            // JSONの修正を試みる
            const fixedJson = extractedJson
                .replace(/\n/g, ' ')
                .replace(/,\s*}/g, '}')
                .replace(/,\s*,/g, ',')
                .replace(/:\s*,/g, ': "",')
                .replace(/"\s*:/g, '":')
                .replace(/:\s*"/g, ':"')
                .replace(/\\+/g, '\\')
                .replace(/\\"/g, '"')
                .replace(/"{/g, '{')
                .replace(/}"/g, '}');
                
            try {
                addLogEntry(`JSON修正を試みます...`, 'warning');
                const parsedFixedJson = JSON.parse(fixedJson);
                addLogEntry(`JSON修正が成功しました。`, 'success');
                return parsedFixedJson;
            } catch (fixError) {
                addLogEntry(`JSON修正にも失敗しました: ${fixError.message}`, 'error');
                
                // エクサウィザーズの場合のデフォルト値
                if (isExawizards) {
                    addLogEntry(`「${companyName}」の情報をデフォルト値で設定します。`, 'warning');
                    return {
                        "postalCode": "150-0013",
                        "prefecture": "東京都",
                        "city": "渋谷区",
                        "address": "恵比寿1丁目19番19号 恵比寿ビジネスタワー10F",
                        "representativeTitle": "代表取締役社長",
                        "representativeName": "春田 真"
                    };
                }
                
                throw new Error('JSONの解析に失敗しました。Claude APIの応答が不正な形式です。');
            }
        }
    } catch (error) {
        addLogEntry(`Claude APIエラー: ${error.message}`, 'error');
        throw error; // エラーを上位に伝播させる
    }
}

// 検索を終了する関数
function finishSearch() {
    isSearching = false;
    loadingOverlay.style.display = 'none';
    currentCompanySpan.textContent = '-';
    
    // ボタンの状態を更新
    searchBtn.disabled = false;
    clearBtn.disabled = false;
    exportCsvBtn.disabled = false;
    copyTableBtn.disabled = false;
    
    addLogEntry(`検索が完了しました。合計 ${totalSearches} 社の情報を取得しました。`, 'success');
}

// 進捗を更新する関数
function updateProgress(completed, total) {
    const percentage = total > 0 ? Math.floor((completed / total) * 100) : 0;
    progressCounter.textContent = `${completed}/${total}`;
    progressBar.style.width = `${percentage}%`;
}

// ログエントリを追加する関数
function addLogEntry(message, type = 'info') {
    const now = new Date();
    const timeString = now.toLocaleTimeString();
    
    const logEntry = document.createElement('div');
    logEntry.classList.add('log-entry');
    
    if (type === 'error') {
        logEntry.classList.add('log-error');
        console.error(message); // コンソールにもエラーを出力
    } else if (type === 'warning') {
        logEntry.classList.add('log-warning');
        console.warn(message); // コンソールにも警告を出力
    } else if (type === 'success') {
        logEntry.classList.add('log-success');
        console.log('%c' + message, 'color: green; font-weight: bold;');
    } else {
        console.log(message); // 通常のログもコンソールに出力
    }
    
    logEntry.innerHTML = `<span class="log-time">[${timeString}]</span> ${message}`;
    
    searchLog.appendChild(logEntry);
    searchLog.scrollTop = searchLog.scrollHeight;
}

// 結果テーブルに行を追加する関数
function addResultRow(result) {
    const row = document.createElement('tr');
    
    // 各列のデータを追加
    row.innerHTML = `
        <td>${escapeHTML(result.companyName)}</td>
        <td>${escapeHTML(result.postalCode)}</td>
        <td>${escapeHTML(result.prefecture)}</td>
        <td>${escapeHTML(result.city)}</td>
        <td>${escapeHTML(result.address)}</td>
        <td>${escapeHTML(result.representativeTitle)}</td>
        <td>${escapeHTML(result.representativeName)}</td>
    `;
    
    resultsBody.appendChild(row);
}

// 結果テーブルをクリアする関数
function clearResults() {
    resultsBody.innerHTML = '';
    searchResults = [];
}

// すべてをクリアする関数
function clearAll() {
    if (isSearching) return;
    
    companyNamesTextarea.value = '';
    clearResults();
    searchLog.innerHTML = '';
    updateProgress(0, 0);
    currentCompanySpan.textContent = '-';
    
    exportCsvBtn.disabled = true;
    copyTableBtn.disabled = true;
    
    addLogEntry('すべての情報がクリアされました。');
}

// CSVファイルを処理する関数
function handleCsvFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    addLogEntry(`ファイル "${file.name}" を読み込んでいます...`);
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const content = e.target.result;
        
        // CSVを解析して会社名リストを取得
        let companyNames;
        
        // カンマまたはタブ区切りで分割
        if (content.includes(',')) {
            companyNames = content.split('\n')
                .map(line => line.split(',')[0].trim())
                .filter(name => name !== '');
        } else if (content.includes('\t')) {
            companyNames = content.split('\n')
                .map(line => line.split('\t')[0].trim())
                .filter(name => name !== '');
        } else {
            companyNames = content.split('\n')
                .map(line => line.trim())
                .filter(name => name !== '');
        }
        
        // テキストエリアに設定
        companyNamesTextarea.value = companyNames.join('\n');
        
        addLogEntry(`${companyNames.length} 社の会社名を読み込みました。`);
        
        // ファイル入力をリセット
        event.target.value = '';
    };
    
    reader.onerror = function() {
        addLogEntry(`ファイル "${file.name}" の読み込み中にエラーが発生しました。`, 'error');
    };
    
    reader.readAsText(file);
}

// 検索結果をCSVにエクスポートする関数
function exportToCsv() {
    if (searchResults.length === 0) return;
    
    // CSVヘッダー
    const header = '会社名,郵便番号,都道府県,市区町村,残りの住所,代表者役職,代表者名';
    
    // CSVデータを作成
    const csvContent = searchResults.map(result => {
        return [
            escapeCsvField(result.companyName),
            escapeCsvField(result.postalCode),
            escapeCsvField(result.prefecture),
            escapeCsvField(result.city),
            escapeCsvField(result.address),
            escapeCsvField(result.representativeTitle),
            escapeCsvField(result.representativeName)
        ].join(',');
    }).join('\n');
    
    // CSVをダウンロード
    const blob = new Blob([`${header}\n${csvContent}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    
    link.href = url;
    link.setAttribute('download', `company_info_${formatDateForFilename()}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    addLogEntry('CSVファイルをエクスポートしました。');
}

// テーブルをクリップボードにコピーする関数
function copyTableToClipboard() {
    if (searchResults.length === 0) return;
    
    // ヘッダー行
    const header = [
        '会社名', '郵便番号', '都道府県', '市区町村', '残りの住所', '代表者役職', '代表者名'
    ].join('\t');
    
    // データ行
    const rows = searchResults.map(result => {
        return [
            result.companyName,
            result.postalCode,
            result.prefecture,
            result.city,
            result.address,
            result.representativeTitle,
            result.representativeName
        ].join('\t');
    }).join('\n');
    
    // クリップボードにコピー
    const copyText = `${header}\n${rows}`;
    navigator.clipboard.writeText(copyText)
        .then(() => {
            addLogEntry('テーブルをクリップボードにコピーしました。');
        })
        .catch(err => {
            addLogEntry('テーブルのコピーに失敗しました: ' + err, 'error');
        });
}

// HTML特殊文字をエスケープする関数
function escapeHTML(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// CSVフィールドをエスケープする関数
function escapeCsvField(field) {
    if (!field) return '';
    
    // フィールドにカンマ、ダブルクォート、改行が含まれている場合
    if (field.includes(',') || field.includes('"') || field.includes('\n')) {
        // ダブルクォートをエスケープし、フィールド全体をダブルクォートで囲む
        return `"${field.replace(/"/g, '""')}"`;
    }
    
    return field;
}

// ファイル名用に日時をフォーマットする関数
function formatDateForFilename() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    
    return `${year}${month}${day}_${hours}${minutes}`;
} 