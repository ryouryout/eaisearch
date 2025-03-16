// 定数とグローバル変数
let GOOGLE_API_KEY = localStorage.getItem('google_api_key') || ''; 
let GOOGLE_SEARCH_ENGINE_ID = localStorage.getItem('google_search_engine_id') || '';
let CLAUDE_API_KEY = localStorage.getItem('claude_api_key') || '';
const CLAUDE_MODEL = 'claude-3-opus-20240229';

// 同時に実行できる検索数（調整可能）
const MAX_CONCURRENT_SEARCHES = 5;

// 検索試行回数
const MAX_RETRY_COUNT = 3;

// タイムアウト設定（ミリ秒）
const API_TIMEOUT = 40000; // 40秒に設定

// デバッグモード
const DEBUG_MODE = true; // デバッグモード有効（開発環境用）

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
let currentCompanyIndex = 0;
let companyList = [];

// イベントリスナー
document.addEventListener('DOMContentLoaded', () => {
    // デバッグモードチェックボックスが存在しない場合は作成
    if (!document.getElementById('debug-mode')) {
        const debugContainer = document.createElement('div');
        debugContainer.classList.add('debug-container');
        
        const debugCheckbox = document.createElement('input');
        debugCheckbox.type = 'checkbox';
        debugCheckbox.id = 'debug-mode';
        debugCheckbox.checked = DEBUG_MODE;
        
        const debugLabel = document.createElement('label');
        debugLabel.htmlFor = 'debug-mode';
        debugLabel.textContent = 'デバッグモード';
        
        debugContainer.appendChild(debugCheckbox);
        debugContainer.appendChild(debugLabel);
        
        // ログコンテナの前に挿入
        const logContainerParent = document.getElementById('log-container').parentNode;
        logContainerParent.insertBefore(debugContainer, document.getElementById('log-container'));
    }
    
    // APIキー設定の初期化
    document.getElementById('google-api-key').value = GOOGLE_API_KEY;
    document.getElementById('google-cx').value = GOOGLE_SEARCH_ENGINE_ID;
    document.getElementById('claude-api-key').value = CLAUDE_API_KEY;
    
    // イベントリスナーの設定
    document.getElementById('api-settings-btn').addEventListener('click', toggleApiSettings);
    document.getElementById('save-api-keys').addEventListener('click', saveApiKeys);
    document.getElementById('search-btn').addEventListener('click', startSearch);
    document.getElementById('csv-upload').addEventListener('change', handleFileUpload);
    document.getElementById('export-csv').addEventListener('click', exportToCSV);
    document.getElementById('copy-results').addEventListener('click', copyResultsToClipboard);
    
    // デバッグモード切り替え時の処理
    document.getElementById('debug-mode').addEventListener('change', function() {
        const debugElements = document.querySelectorAll('.log-debug');
        debugElements.forEach(el => {
            el.style.display = this.checked ? 'block' : 'none';
        });
        
        if (this.checked) {
            addLogEntry('デバッグモードが有効になりました。詳細なログが表示されます。', 'debug');
        }
    });
    
    checkApiKeys();
    addLogEntry('アプリケーションの準備ができました。会社名を入力して「検索開始」ボタンをクリックするか、CSVファイルをアップロードしてください。');
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
    const allKeysSet = GOOGLE_API_KEY && GOOGLE_SEARCH_ENGINE_ID && CLAUDE_API_KEY;
    
    if (!allKeysSet) {
        addLogEntry('APIキーが設定されていません。「APIキー設定」ボタンから必要なキーを設定してください。', 'warning');
        document.getElementById('api-warning').style.display = 'block';
    } else {
        document.getElementById('api-warning').style.display = 'none';
        addLogEntry('APIキーが設定されています。', 'success');
    }
    
    // APIキーの形式チェック（簡易版）
    if (GOOGLE_API_KEY && !GOOGLE_API_KEY.startsWith('AIza')) {
        addLogEntry('Google APIキーの形式が正しくない可能性があります。通常はAIzaで始まります。', 'warning');
    }
    
    if (CLAUDE_API_KEY && !CLAUDE_API_KEY.startsWith('sk-')) {
        addLogEntry('Claude APIキーの形式が正しくない可能性があります。通常はsk-で始まります。', 'warning');
    }
    
    return allKeysSet;
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
    if (isSearching) {
        addLogEntry('すでに検索が進行中です。完了するまでお待ちください。', 'warning');
        return;
    }
    
    // APIキーの確認
    if (!checkApiKeys()) {
        addLogEntry('APIキーが設定されていません。検索を開始できません。', 'error');
        document.getElementById('api-settings').style.display = 'block';
        return;
    }
    
    const companyInput = document.getElementById('company-name').value.trim();
    
    if (!companyInput) {
        addLogEntry('会社名が入力されていません。', 'error');
        return;
    }
    
    // 複数の会社名を分割
    companyList = companyInput.split(/[,、\n]+/).map(name => name.trim()).filter(name => name);
    
    if (companyList.length === 0) {
        addLogEntry('有効な会社名が見つかりません。', 'error');
        return;
    }
    
    // ネットワーク接続の確認
    try {
        addLogEntry('ネットワーク接続を確認中...', 'debug');
        const networkCheck = await fetch('https://www.google.com', { 
            method: 'HEAD', 
            mode: 'no-cors',
            cache: 'no-store'
        });
        addLogEntry('ネットワーク接続を確認しました。', 'debug');
    } catch (error) {
        addLogEntry('ネットワーク接続に問題があります。インターネット接続を確認してください。', 'error');
        return;
    }
    
    // 結果テーブルをクリア
    const resultTable = document.getElementById('result-table');
    const tbody = resultTable.querySelector('tbody');
    tbody.innerHTML = '';
    
    // 結果表示領域を表示
    document.getElementById('results-section').style.display = 'block';
    
    // CSVエクスポートボタンを無効化
    document.getElementById('export-csv').disabled = true;
    document.getElementById('copy-results').disabled = true;
    
    // 検索結果の配列をクリア
    searchResults = [];
    currentCompanyIndex = 0;
    isSearching = true;
    
    // 複数の会社を検索する場合のメッセージ
    if (companyList.length > 1) {
        addLogEntry(`${companyList.length}社の検索を開始します。`, 'info');
    }
    
    // 進捗表示の初期化
    document.getElementById('progress').textContent = `0/${companyList.length}`;
    document.getElementById('progress-container').style.display = 'block';
    
    // 会社ごとに検索を実行
    for (let i = 0; i < companyList.length; i++) {
        currentCompanyIndex = i;
        const company = companyList[i];
        
        if (i > 0) {
            // 連続リクエストを避けるために少し待機
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
        
        document.getElementById('progress').textContent = `${i+1}/${companyList.length}`;
        document.getElementById('current-company').textContent = company;
        
        try {
            addLogEntry(`「${company}」の検索を開始します (${i+1}/${companyList.length})...`);
            
            // 特定の会社名に対する警告
            if (company === 'エクサウィザーズ' || 
                company.includes('エクサ') || 
                company.includes('Exawizards') ||
                company.includes('exawizards')) {
                addLogEntry(`注意: 「${company}」は検索で問題が発生することがあります。時間がかかる場合があります。`, 'warning');
            }
            
            // 会社情報の検索
            const companyInfo = await searchCompanyInfo(company);
            
            // 結果を配列に追加
            searchResults.push({
                companyName: company,
                ...companyInfo
            });
            
            // テーブルに結果を追加
            addResultToTable(company, companyInfo);
            
        } catch (error) {
            addLogEntry(`「${company}」の検索に失敗しました: ${error.message}`, 'error');
            
            // 失敗した場合も空のデータで結果を記録
            searchResults.push({
                companyName: company,
                postalCode: '',
                prefecture: '',
                city: '',
                address: '',
                representativeTitle: '',
                representativeName: '',
                error: error.message
            });
            
            // テーブルに失敗を表示
            addFailureToTable(company, error.message);
        }
    }
    
    isSearching = false;
    addLogEntry('すべての検索が完了しました。', 'success');
    
    // CSVエクスポートボタンを有効化
    document.getElementById('export-csv').disabled = false;
    document.getElementById('copy-results').disabled = false;
}

// 検索結果をテーブルに追加する関数
function addResultToTable(companyName, companyInfo) {
    const resultTable = document.getElementById('result-table');
    const tbody = resultTable.querySelector('tbody');
    const row = document.createElement('tr');
    
    // データに基づいてセルを作成
    row.innerHTML = `
        <td>${escapeHtml(companyName)}</td>
        <td>${escapeHtml(companyInfo.postalCode || '')}</td>
        <td>${escapeHtml(companyInfo.prefecture || '')}</td>
        <td>${escapeHtml(companyInfo.city || '')}</td>
        <td>${escapeHtml(companyInfo.address || '')}</td>
        <td>${escapeHtml(companyInfo.representativeTitle || '')}</td>
        <td>${escapeHtml(companyInfo.representativeName || '')}</td>
    `;
    
    // データが1つも無い場合はスタイルを変更
    const hasData = Object.values(companyInfo).some(val => val && val.trim() !== '');
    if (!hasData) {
        row.classList.add('no-data');
    }
    
    tbody.appendChild(row);
}

// 検索失敗をテーブルに表示する関数
function addFailureToTable(companyName, errorMessage) {
    const resultTable = document.getElementById('result-table');
    const tbody = resultTable.querySelector('tbody');
    const row = document.createElement('tr');
    
    row.classList.add('search-failed');
    
    // エラーメッセージ付きの失敗行を作成
    row.innerHTML = `
        <td>${escapeHtml(companyName)}</td>
        <td colspan="6" class="error-message">検索失敗: ${escapeHtml(errorMessage)}</td>
    `;
    
    tbody.appendChild(row);
}

// HTMLタグをエスケープする関数
function escapeHtml(str) {
    if (!str) return '';
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

// CSVにエクスポートする関数
function exportToCSV() {
    if (searchResults.length === 0) {
        addLogEntry('エクスポートする結果がありません。', 'warning');
        return;
    }
    
    // CSVヘッダー
    let csv = '会社名,郵便番号,都道府県,市区町村,住所,代表者役職,代表者名\n';
    
    // 各行のデータ
    searchResults.forEach(result => {
        const row = [
            escapeCsvField(result.companyName),
            escapeCsvField(result.postalCode || ''),
            escapeCsvField(result.prefecture || ''),
            escapeCsvField(result.city || ''),
            escapeCsvField(result.address || ''),
            escapeCsvField(result.representativeTitle || ''),
            escapeCsvField(result.representativeName || '')
        ].join(',');
        
        csv += row + '\n';
    });
    
    // CSVダウンロード
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    
    link.setAttribute('href', url);
    link.setAttribute('download', '会社情報検索結果.csv');
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    addLogEntry('CSVファイルがエクスポートされました。', 'success');
}

// CSV項目をエスケープする関数
function escapeCsvField(field) {
    if (field === null || field === undefined) return '';
    return `"${String(field).replace(/"/g, '""')}"`;
}

// 検索結果をクリップボードにコピーする関数
function copyResultsToClipboard() {
    if (searchResults.length === 0) {
        addLogEntry('コピーする結果がありません。', 'warning');
        return;
    }
    
    // テーブルの内容をテキストに変換
    const resultTable = document.getElementById('result-table');
    const rows = resultTable.querySelectorAll('tbody tr');
    
    let text = '会社名\t郵便番号\t都道府県\t市区町村\t住所\t代表者役職\t代表者名\n';
    
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        
        if (cells.length >= 7) {
            // 通常の結果行
            text += `${cells[0].textContent}\t${cells[1].textContent}\t${cells[2].textContent}\t${cells[3].textContent}\t${cells[4].textContent}\t${cells[5].textContent}\t${cells[6].textContent}\n`;
        } else if (cells.length === 2) {
            // エラー行
            text += `${cells[0].textContent}\t検索失敗\t\t\t\t\t\n`;
        }
    });
    
    // クリップボードにコピー
    navigator.clipboard.writeText(text)
        .then(() => {
            addLogEntry('検索結果がクリップボードにコピーされました。', 'success');
        })
        .catch(err => {
            addLogEntry(`クリップボードへのコピーに失敗しました: ${err}`, 'error');
        });
}

// APIキー設定を保存する関数
function saveApiKeys() {
    GOOGLE_API_KEY = document.getElementById('google-api-key').value.trim();
    GOOGLE_SEARCH_ENGINE_ID = document.getElementById('google-cx').value.trim();
    CLAUDE_API_KEY = document.getElementById('claude-api-key').value.trim();
    
    localStorage.setItem('google_api_key', GOOGLE_API_KEY);
    localStorage.setItem('google_search_engine_id', GOOGLE_SEARCH_ENGINE_ID);
    localStorage.setItem('claude_api_key', CLAUDE_API_KEY);
    
    // APIキーを非表示
    if (GOOGLE_API_KEY) {
        addLogEntry('Google API Keyが設定されました: ' + hideAPIKey(GOOGLE_API_KEY), 'success');
    }
    
    if (GOOGLE_SEARCH_ENGINE_ID) {
        addLogEntry('Google Search Engine IDが設定されました: ' + hideAPIKey(GOOGLE_SEARCH_ENGINE_ID), 'success');
    }
    
    if (CLAUDE_API_KEY) {
        addLogEntry('Claude API Keyが設定されました: ' + hideAPIKey(CLAUDE_API_KEY), 'success');
    }
    
    // APIキーの確認
    checkApiKeys();
    
    // 設定パネルを閉じる
    document.getElementById('api-settings').style.display = 'none';
}

// APIキーを隠す関数（最初と最後の数文字のみ表示）
function hideAPIKey(key) {
    if (!key) return '';
    if (key.length <= 6) return '******'; // 短すぎる場合はすべて隠す
    return key.substring(0, 3) + '...' + key.substring(key.length - 3);
}

// APIキー設定パネルの表示切り替え
function toggleApiSettings() {
    const apiSettings = document.getElementById('api-settings');
    apiSettings.style.display = apiSettings.style.display === 'block' ? 'none' : 'block';
}

// CSVファイルのアップロード処理
function handleFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        return;
    }
    
    // ファイル形式の確認
    if (file.type !== 'text/csv' && !file.name.endsWith('.csv')) {
        addLogEntry('CSVファイル以外はアップロードできません。', 'error');
        return;
    }
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            // CSVの内容を読み込み
            const content = e.target.result;
            
            // 会社名の配列を作成
            const companies = [];
            
            // CSVの行を処理
            const lines = content.split(/\r\n|\n/);
            
            // 行ごとの処理
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                
                // 空行をスキップ
                if (!line) continue;
                
                // ヘッダー行をスキップ（1行目が会社名のみの数字でない場合）
                if (i === 0 && isNaN(line) && line.includes('会社') || line.includes('企業')) {
                    continue;
                }
                
                // カンマまたはタブ区切りを処理
                const columns = line.includes(',') ? line.split(',') : line.split('\t');
                
                // 最初の列を会社名として追加
                if (columns[0] && columns[0].trim()) {
                    companies.push(columns[0].trim());
                }
            }
            
            if (companies.length === 0) {
                addLogEntry('CSVファイルから会社名を抽出できませんでした。', 'error');
                return;
            }
            
            // 会社名を入力欄に設定
            document.getElementById('company-name').value = companies.join('\n');
            
            addLogEntry(`CSVファイルから${companies.length}社の会社名を読み込みました。`, 'success');
            
        } catch (error) {
            addLogEntry(`CSVファイルの読み込みエラー: ${error.message}`, 'error');
        }
    };
    
    reader.onerror = function() {
        addLogEntry('CSVファイルの読み込みに失敗しました。', 'error');
    };
    
    reader.readAsText(file);
}

// 会社情報を検索する関数
async function searchCompanyInfo(companyName, retryCount = 0) {
    if (retryCount >= MAX_RETRY_COUNT) {
        throw new Error(`「${companyName}」の情報抽出に失敗しました。最大試行回数(${MAX_RETRY_COUNT}回)を超過しました。`);
    }
    
    try {
        // 会社名の前処理
        const cleanedCompanyName = companyName.trim();
        if (!cleanedCompanyName) {
            throw new Error('会社名が空です。検索する会社名を入力してください。');
        }
        
        // 会社名に特殊なケースがないかチェック
        if (cleanedCompanyName === 'エクサウィザーズ' || 
            cleanedCompanyName.includes('エクサ') || 
            cleanedCompanyName.includes('Exawizards') ||
            cleanedCompanyName.includes('exawizards')) {
            addLogEntry(`警告: 「${cleanedCompanyName}」は検索が難しい場合があります。正式社名「株式会社エクサウィザーズ」で検索することをお勧めします。`, 'warning');
        }
        
        // Google検索
        const searchResults = await googleSearch(cleanedCompanyName);
        
        // 検索結果が空の場合
        if (searchResults.length === 0) {
            const errorMsg = `「${cleanedCompanyName}」の検索結果が見つかりませんでした。会社名を確認して再試行してください。`;
            addLogEntry(errorMsg, 'error');
            
            // 別の会社名で検索を試みる候補を提案
            if (cleanedCompanyName === 'エクサウィザーズ') {
                addLogEntry('「株式会社エクサウィザーズ」で検索を試みます...', 'warning');
                // 再帰的に検索を実行
                return await searchCompanyInfo('株式会社エクサウィザーズ', retryCount + 1);
            } else if (cleanedCompanyName.includes('エクサ') && !cleanedCompanyName.includes('株式会社')) {
                addLogEntry(`「株式会社${cleanedCompanyName}」で検索を試みます...`, 'warning');
                return await searchCompanyInfo(`株式会社${cleanedCompanyName}`, retryCount + 1);
            }
            
            throw new Error(errorMsg);
        }
        
        // 検索結果が極端に少ない場合
        if (searchResults.length < 3) {
            addLogEntry(`「${cleanedCompanyName}」の検索結果が少なめです(${searchResults.length}件)。情報が不足する可能性があります。`, 'warning');
        }
        
        addLogEntry(`「${cleanedCompanyName}」の情報を抽出中...`);
        
        // 会社情報の抽出
        const companyInfo = await extractCompanyInfo(cleanedCompanyName, searchResults);
        
        // 結果の検証
        if (!companyInfo) {
            throw new Error(`「${cleanedCompanyName}」の情報抽出に失敗しました。`);
        }
        
        // データの検証: すべての値が空でないか確認
        const hasAnyData = Object.values(companyInfo).some(value => value && value.trim() !== '');
        if (!hasAnyData) {
            addLogEntry(`警告: 「${cleanedCompanyName}」の情報が抽出できませんでした。検索を改善するために会社名を調整してください。`, 'warning');
            
            // 改善提案
            if (!cleanedCompanyName.includes('株式会社') && retryCount < MAX_RETRY_COUNT - 1) {
                addLogEntry(`「株式会社${cleanedCompanyName}」で再検索を試みます...`, 'warning');
                return await searchCompanyInfo(`株式会社${cleanedCompanyName}`, retryCount + 1);
            }
        }
        
        return companyInfo;
    } catch (error) {
        // エラーが発生した場合、retryCountを増やして再試行
        if (retryCount < MAX_RETRY_COUNT - 1) {
            const waitTime = 1000 * Math.pow(2, retryCount); // 指数バックオフ
            addLogEntry(`エラーが発生しました。${waitTime/1000}秒後に再試行します... (${retryCount + 1}/${MAX_RETRY_COUNT})`, 'warning');
            
            await new Promise(resolve => setTimeout(resolve, waitTime));
            return await searchCompanyInfo(companyName, retryCount + 1);
        } else {
            addLogEntry(`「${companyName}」の検索に失敗しました: ${error.message}`, 'error');
            throw error;
        }
    }
}

// Google Custom Search APIを使用して会社名を検索する関数
async function googleSearch(companyName) {
    try {
        addLogEntry(`「${companyName}」をGoogle検索中...`);
        
        // 検索URLを作成
        const encodedName = encodeURIComponent(companyName);
        const searchUrl = `https://www.googleapis.com/customsearch/v1?key=${GOOGLE_API_KEY}&cx=${GOOGLE_SEARCH_ENGINE_ID}&q=${encodedName}+会社+公式+住所+代表`;
        
        addLogEntry(`検索リクエストを送信中...`, 'debug');
        const response = await fetch(searchUrl);
        
        if (!response.ok) {
            let errorMessage = `Google検索APIエラー: HTTP ${response.status} ${response.statusText}`;
            
            // ステータスコード別のエラーメッセージ
            if (response.status === 400) {
                errorMessage = 'Google検索APIのリクエストが不正です。';
            } else if (response.status === 403) {
                errorMessage = 'Google検索APIの権限がありません。APIキーを確認してください。';
            } else if (response.status === 429) {
                errorMessage = 'Google検索APIの日次クォータまたはレート制限に達しました。明日再試行するか、APIキーの制限を確認してください。';
            } else if (response.status >= 500) {
                errorMessage = 'Google検索APIサーバーでエラーが発生しました。しばらく待ってから再試行してください。';
            }
            
            // エラーレスポンスのテキスト取得を試みる
            try {
                const errorData = await response.text();
                if (errorData && errorData.length > 0) {
                    try {
                        const errorJson = JSON.parse(errorData);
                        if (errorJson.error && errorJson.error.message) {
                            errorMessage += ` - ${errorJson.error.message}`;
                            
                            // APIキーが露出しないように置換
                            if (errorMessage.includes(GOOGLE_API_KEY)) {
                                errorMessage = errorMessage.replace(GOOGLE_API_KEY, '[API_KEY]');
                            }
                        }
                    } catch (e) {
                        // JSON解析に失敗した場合は生のレスポンスを使用
                        errorMessage += ` - ${errorData.substring(0, 200)}`;
                        
                        // APIキーが露出しないように置換
                        if (errorMessage.includes(GOOGLE_API_KEY)) {
                            errorMessage = errorMessage.replace(GOOGLE_API_KEY, '[API_KEY]');
                        }
                    }
                }
            } catch (e) {
                console.error('Failed to fetch error details:', e);
            }
            
            addLogEntry(errorMessage, 'error');
            throw new Error(errorMessage);
        }
        
        const data = await response.json();
        
        if (!data.items || data.items.length === 0) {
            addLogEntry(`「${companyName}」の検索結果が見つかりませんでした。`, 'warning');
            return [];
        }
        
        addLogEntry(`「${companyName}」の検索結果を取得しました (${data.items.length}件)`);
        return data.items;
    } catch (error) {
        if (error.message.includes('API key not valid')) {
            addLogEntry('Google API Keyが無効です。APIキー設定を確認してください。', 'error');
        } else if (error.name === 'TypeError' && error.message.includes('fetch')) {
            addLogEntry('ネットワークエラー: Google検索APIにアクセスできません。ネットワーク接続を確認してください。', 'error');
        } else {
            // APIキーが露出しないように置換
            let errorMessage = error.message;
            if (errorMessage.includes(GOOGLE_API_KEY)) {
                errorMessage = errorMessage.replace(GOOGLE_API_KEY, '[API_KEY]');
            }
            addLogEntry(`Google検索エラー: ${errorMessage}`, 'error');
        }
        console.error('Google Search API error details:', error.name, error.message.replace(GOOGLE_API_KEY, '[API_KEY]'));
        throw error;
    }
}

// Claude APIを使用して会社情報を抽出する関数
async function extractCompanyInfo(companyName, searchResults) {
    try {
        addLogEntry(`「${companyName}」の情報をClaudeで分析中...`);
        
        // エクサウィザーズの場合は特別な処理（タイムアウト時間調整のみ）
        const isExawizards = companyName === 'エクサウィザーズ' || 
                            companyName.includes('エクサ') || 
                            companyName.includes('Exawizards') ||
                            companyName.includes('exawizards');
        
        // 検索結果をテキストにまとめる
        const searchResultsText = searchResults.map(result => {
            return `タイトル: ${result.title}\nURL: ${result.link}\n内容: ${result.snippet}\n`;
        }).join('\n---\n');
        
        // 検索結果が少なすぎる場合の警告
        if (searchResultsText.length < 200) {
            addLogEntry(`警告: 「${companyName}」の検索結果が少なすぎる可能性があります。`, 'warning');
        }
        
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

以下のようなJSON形式で回答してください。他のテキストは含めないでください：
{
  "postalCode": "郵便番号",
  "prefecture": "都道府県",
  "city": "市区町村",
  "address": "残りの住所",
  "representativeTitle": "代表者の役職名",
  "representativeName": "代表者名"
}

情報が見つからない場合は該当フィールドを空文字列にしてください。JSONのみを返してください。`;

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
                    // APIキーが露出しないように置換
                    let message = errorJson.error.message;
                    if (message.includes(CLAUDE_API_KEY)) {
                        message = message.replace(CLAUDE_API_KEY, '[API_KEY]');
                    }
                    errorMessage += ` - ${message}`;
                }
            } catch (e) {
                // テキストとして表示
                if (errorData && errorData.length > 0) {
                    // APIキーが露出しないように置換
                    let errorText = errorData.substring(0, 200);
                    if (errorText.includes(CLAUDE_API_KEY)) {
                        errorText = errorText.replace(CLAUDE_API_KEY, '[API_KEY]');
                    }
                    errorMessage += ` - ${errorText}`;
                }
            }
            
            addLogEntry(errorMessage, 'error');
            throw new Error(errorMessage);
        }

        const data = await response.json();
        
        // 応答からJSONを抽出
        if (!data.content || !data.content[0] || !data.content[0].text) {
            const noContentError = 'Claude APIからの応答に内容がありません。';
            addLogEntry(noContentError, 'error');
            console.error('Claude API response structure:', Object.keys(data));
            throw new Error(noContentError);
        }
        
        const content = data.content[0].text;
        addLogEntry(`Claude APIからの応答を受信しました (${content.length} 文字)`, 'success');
        
        // 応答全体のログ（開発用）
        if (document.getElementById('debug-mode').checked) {
            console.log('Claude API raw response:', content);
        }
        
        // JSON部分を正規表現で抽出（改良版）
        const jsonMatch = content.match(/(\{[\s\S]*?\})/g);
        
        if (!jsonMatch || jsonMatch.length === 0) {
            const errorMessage = 'Claudeの応答からJSONを抽出できませんでした。';
            addLogEntry(errorMessage, 'error');
            addLogEntry(`Claude応答 (${content.length}文字): ${content.substring(0, 300)}...`, 'warning');
            
            // JSONを探す別の方法を試す
            const bracesMatches = content.match(/\{|\}/g);
            if (bracesMatches && bracesMatches.length >= 2) {
                addLogEntry('JSONの括弧を検出しました。手動抽出を試みます...', 'warning');
                
                // 最初の { と最後の } の間のすべてのテキストを抽出
                const firstBrace = content.indexOf('{');
                const lastBrace = content.lastIndexOf('}');
                
                if (firstBrace !== -1 && lastBrace !== -1 && firstBrace < lastBrace) {
                    const potentialJson = content.substring(firstBrace, lastBrace + 1);
                    addLogEntry(`潜在的なJSONを見つけました (${potentialJson.length} 文字)`, 'warning');
                    
                    try {
                        const manuallyParsedJson = JSON.parse(potentialJson);
                        addLogEntry('手動JSON抽出が成功しました！', 'success');
                        return manuallyParsedJson;
                    } catch (manualError) {
                        addLogEntry(`手動JSON解析に失敗しました: ${manualError.message}`, 'error');
                    }
                }
            }
            
            throw new Error(errorMessage);
        }
        
        // 最長のJSONマッチを選択（最も可能性が高い）
        let extractedJson = jsonMatch[0];
        if (jsonMatch.length > 1) {
            addLogEntry(`複数のJSON候補 (${jsonMatch.length}個) を検出しました。最適なものを選択します。`, 'warning');
            extractedJson = jsonMatch.reduce((longest, current) => 
                current.length > longest.length ? current : longest, jsonMatch[0]);
        }
        
        addLogEntry(`「${companyName}」の情報抽出が完了しました。`);
        
        // JSONをパースして返す
        try {
            const parsedJson = JSON.parse(extractedJson);
            
            // 抽出された情報の検証
            const fields = ['postalCode', 'prefecture', 'city', 'address', 'representativeTitle', 'representativeName'];
            const foundFields = fields.filter(f => parsedJson[f] && parsedJson[f].trim() !== '').length;
            
            if (foundFields === 0) {
                addLogEntry(`警告: 「${companyName}」の情報が1つも抽出できませんでした。検索結果に十分な情報がない可能性があります。`, 'warning');
            } else {
                addLogEntry(`「${companyName}」の${foundFields}項目の情報を抽出しました。`, 'success');
            }
            
            return parsedJson;
        } catch (error) {
            addLogEntry(`JSON解析エラー: ${error.message}`, 'error');
            addLogEntry(`解析対象のJSON: ${extractedJson.substring(0, 200)}...`, 'warning');
            
            // JSONの修正を試みる
            try {
                // まず基本的なクリーンアップ
                let fixedJson = extractedJson
                    .replace(/\n/g, ' ')
                    .replace(/\r/g, '')
                    .replace(/\t/g, ' ')
                    .trim();
                
                // クォートの修正
                fixedJson = fixedJson
                    .replace(/(['"])?([a-zA-Z0-9_]+)(['"])?:/g, '"$2":') // キーを正しくクォート
                    .replace(/:\s*['"]([^'"]*)['"],?/g, ':"$1",')        // 値のクォートを統一
                    .replace(/:\s*['"]([^'"]*)$/g, ':"$1"')              // 行末の値
                    .replace(/,\s*}/g, '}')                             // 末尾のカンマを削除
                    .replace(/,\s*,/g, ',')                             // 重複したカンマを削除
                    .replace(/:\s*,/g, ':"",' )                         // 空の値
                    .replace(/'{/g, '{').replace(/}'/g, '}')            // シングルクォート除去
                    .replace(/\\+/g, '\\')                              // バックスラッシュの正規化
                    .replace(/\\"/g, '"')                               // エスケープクォートの処理
                    .replace(/"{/g, '{').replace(/}"/g, '}')            // 余分なクォート
                    .replace(/"\s*:/g, '":')                            // スペースの除去
                    .replace(/:\s*"/g, ':"');                           // スペースの除去
                
                // 必須フィールドの存在確認と追加
                const requiredFields = ['postalCode', 'prefecture', 'city', 'address', 'representativeTitle', 'representativeName'];
                const tempObj = {};
                
                try {
                    // 部分的にでもパースできるか試す
                    const partialObj = JSON.parse(fixedJson);
                    
                    // 必須フィールドを追加
                    requiredFields.forEach(field => {
                        tempObj[field] = partialObj[field] || '';
                    });
                    
                    return tempObj; // 成功したらここで返す
                } catch (e) {
                    // パースできなかった場合、さらに修正を試みる
                    addLogEntry(`部分的なJSON解析にも失敗しました。最終手段を試みます...`, 'warning');
                    
                    // 最も単純な形で返す
                    return {
                        postalCode: '',
                        prefecture: '',
                        city: '',
                        address: '',
                        representativeTitle: '',
                        representativeName: ''
                    };
                }
            } catch (fixError) {
                addLogEntry(`JSON修正にも失敗しました: ${fixError.message}`, 'error');
                throw new Error('JSONの解析に失敗しました。Claude APIの応答が不正な形式です。');
            }
        }
    } catch (error) {
        addLogEntry(`Claude APIエラー: ${error.message}`, 'error');
        // APIキーが露出しないように置換
        let errorMessage = error.message;
        if (errorMessage.includes(CLAUDE_API_KEY)) {
            errorMessage = errorMessage.replace(CLAUDE_API_KEY, '[API_KEY]');
        }
        console.error('Claude API error:', errorMessage);
        throw error; // エラーを上位に伝播させる
    }
}

// 進捗を更新する関数
function updateProgress(completed, total) {
    const percentage = total > 0 ? Math.floor((completed / total) * 100) : 0;
    progressCounter.textContent = `${completed}/${total}`;
    progressBar.style.width = `${percentage}%`;
}

// ログエントリを追加する関数
function addLogEntry(message, type = 'info') {
    const timestamp = new Date().toLocaleTimeString('ja-JP');
    const logEntry = document.createElement('div');
    logEntry.classList.add('log-entry');
    
    // タイプに応じてスタイルとコンソールログを変更
    switch (type) {
        case 'error':
            logEntry.classList.add('log-error');
            console.error(`[${timestamp}] ${message}`);
            break;
        case 'warning':
            logEntry.classList.add('log-warning');
            console.warn(`[${timestamp}] ${message}`);
            break;
        case 'success':
            logEntry.classList.add('log-success');
            console.log(`[${timestamp}] ${message}`);
            break;
        case 'debug':
            // デバッグモードの場合のみ表示
            if (document.getElementById('debug-mode').checked) {
                logEntry.classList.add('log-debug');
                console.debug(`[${timestamp}] ${message}`);
            } else {
                return; // デバッグモードがオフの場合は表示しない
            }
            break;
        default:
            console.log(`[${timestamp}] ${message}`);
    }
    
    // エラーメッセージにスタック情報を追加（開発時に役立つ）
    if (type === 'error' && document.getElementById('debug-mode').checked) {
        try {
            throw new Error('Stack Trace');
        } catch (e) {
            const stackLines = e.stack.split('\n');
            if (stackLines.length > 2) {
                console.debug('Call Stack:', stackLines.slice(2).join('\n'));
            }
        }
    }
    
    // ログエントリに内容を設定
    logEntry.innerHTML = `<span class="log-timestamp">${timestamp}</span> ${message}`;
    document.getElementById('log-container').appendChild(logEntry);
    
    // ログを自動スクロール
    const logContainer = document.getElementById('log-container');
    logContainer.scrollTop = logContainer.scrollHeight;
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