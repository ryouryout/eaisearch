// 定数とグローバル変数
const GOOGLE_API_KEY = 'YOUR_GOOGLE_API_KEY'; // 実際の使用時には設定してください
const GOOGLE_SEARCH_ENGINE_ID = 'YOUR_SEARCH_ENGINE_ID'; // 実際の使用時には設定してください
const CLAUDE_API_KEY = 'YOUR_CLAUDE_API_KEY'; // 実際の使用時には設定してください
const CLAUDE_MODEL = 'claude-3-7-sonnet-20250219';

// 同時に実行できる検索数（調整可能）
const MAX_CONCURRENT_SEARCHES = 5;

// 検索試行回数
const MAX_RETRY_COUNT = 2;

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

// 検索を開始する関数
async function startSearch() {
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
                
                addLogEntry(`「${companyName}」の検索が完了しました。（${completedSearches}/${totalSearches}）`);
            })
            .catch(error => {
                // エラーが発生した場合
                completedSearches++;
                updateProgress(completedSearches, totalSearches);
                
                // エラーログを表示
                addLogEntry(`エラー: 「${companyName}」の検索中にエラーが発生しました: ${error.message}`, 'error');
                
                // エラーでも結果配列に追加（空データ）
                searchResults.push({
                    companyName: companyName,
                    postalCode: '',
                    prefecture: '',
                    city: '',
                    address: '',
                    representativeTitle: '',
                    representativeName: ''
                });
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
        
        // ステップ1: Google検索で会社情報を取得
        const searchTerms = generateSearchTerms(companyName);
        addLogEntry(`検索ワード: "${searchTerms[0]}" を試行中...`);
        
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
                addLogEntry(`「${searchTerm}」での検索中にエラーが発生しました。次の検索ワードを試行します...`, 'warning');
            }
        }
        
        if (!searchSuccess) {
            throw new Error('すべての検索ワードでの検索に失敗しました。');
        }
        
        // ステップ2: Claude APIを使用して会社情報を抽出
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
        // リトライ回数が上限に達していない場合はリトライ
        if (retryCount < MAX_RETRY_COUNT) {
            addLogEntry(`「${companyName}」の検索に失敗しました。リトライします... (${retryCount + 1}/${MAX_RETRY_COUNT})`, 'warning');
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
- 分析には入手可能な情報のみを使用してください。不明な情報を推測しないでください。`;

        // Claude APIにリクエストを送信
        const response = await fetch('https://api.anthropic.com/v1/messages', {
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
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Claude APIエラー: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const data = await response.json();
        
        // 応答からJSONを抽出
        const content = data.content[0].text;
        const jsonMatch = content.match(/\{[\s\S]*\}/);
        
        if (!jsonMatch) {
            throw new Error('Claudeの応答からJSONを抽出できませんでした。');
        }
        
        const extractedJson = jsonMatch[0];
        addLogEntry(`「${companyName}」の情報抽出が完了しました。`);
        
        // JSONをパースして返す
        return JSON.parse(extractedJson);
    } catch (error) {
        console.error('Claude APIエラー:', error);
        addLogEntry(`「${companyName}」の情報抽出中にエラーが発生しました。`, 'error');
        
        // エラーの場合は空のオブジェクトを返す
        return {
            postalCode: '',
            prefecture: '',
            city: '',
            address: '',
            representativeTitle: '',
            representativeName: ''
        };
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
    } else if (type === 'warning') {
        logEntry.classList.add('log-warning');
    } else if (type === 'success') {
        logEntry.classList.add('log-success');
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