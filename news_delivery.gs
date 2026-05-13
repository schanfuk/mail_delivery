// ──────────────────────────────────────────────────────────────
//  福岡学生情報メルマガ - 配信スクリプト
// ──────────────────────────────────────────────────────────────
//  初回セットアップ:
//    1. setupBucketName() を一度実行して GCS バケット名を保存
//       (GCSバックアップを使わない場合は空欄でOK)
//    2. fetchAndStoreNews を時間トリガーに登録
// ──────────────────────────────────────────────────────────────

// 各カテゴリの採用件数を「ご当地枠 / 全国区枠」に分割
const CATEGORY_LIMITS = {
  "お得情報・クーポン":       { local: 2, national: 3 },
  "ハッカソン・コンテスト":   { local: 3, national: 1 },
  "留学・海外プログラム":     { local: 2, national: 2 },
  "インターン・プログラム":   { local: 3, national: 2 }
};

const PLUS_WORDS = {
  "お得情報・クーポン": [
    "学割", "クーポン", "割引", "無料", "半額", "OFF", "キャンペーン",
    "限定", "プレゼント", "ポイント", "セール", "特典"
  ],
  "ハッカソン・コンテスト": [
    "ハッカソン", "ビジコン", "ビジネスコンテスト", "コンテスト", "コンペ",
    "ピッチ", "賞金", "アイデアソン", "スタートアップ", "ピッチイベント"
  ],
  "留学・海外プログラム": [
    "留学", "奨学金", "JASSO", "Fulbright", "交換留学", "海外派遣",
    "海外研修", "海外プログラム", "フェロー", "給付", "短期留学"
  ],
  "インターン・プログラム": [
    "インターン", "インターンシップ", "募集", "学生プログラム", "ワークショップ",
    "勉強会", "ミートアップ", "メンタリング", "プログラム", "実習", "見学会"
  ]
};

const NATIONAL_BRAND_WORDS = [
  "Amazon", "アマゾン", "Apple", "Microsoft", "Adobe", "Netflix", "Spotify",
  "Google", "YouTube", "Disney", "ユニクロ", "GU", "スタバ", "マック",
  "JR", "ANA", "JAL", "楽天", "Yahoo", "LINE", "メルカリ", "PayPay"
];

const LOCATION_PLUS_WORDS = ["福岡", "博多", "天神", "九州", "Fukuoka"];

const STRONG_WORDS = ["募集中", "募集開始", "締切", "発表", "決定", "新設", "開催"];
const EXCLUDE_WORDS = ["終了", "中止", "延期", "PR", "広告", "詐欺", "まとめ", "ランキング"];

// ──────────────────────────────────────────────────────────────
//  初回セットアップ: GCSバケット名をスクリプトプロパティに保存
//  - GASエディタの関数選択から "setupBucketName" を一度だけ実行
//  - 空欄で保存するとGCSバックアップが無効化される
// ──────────────────────────────────────────────────────────────
function setupBucketName() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty('BUCKET_NAME') || '(未設定)';
  const response = ui.prompt(
    'GCSバケット名の設定',
    `現在の設定: ${current}\n\nGoogle Cloud Storage のバケット名を入力してください。\n(空欄でGCSバックアップを無効化)`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const name = response.getResponseText().trim();
  const props = PropertiesService.getScriptProperties();
  if (name) {
    props.setProperty('BUCKET_NAME', name);
    ui.alert(`保存しました: ${name}`);
  } else {
    props.deleteProperty('BUCKET_NAME');
    ui.alert('GCSバックアップを無効にしました');
  }
}

// ──────────────────────────────────────────────────────────────
//  メイン処理
// ──────────────────────────────────────────────────────────────
function fetchAndStoreNews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('設定');
  if (!configSheet) {
    console.error('「設定」シートが見つかりません');
    return;
  }
  const lastRowConfig = configSheet.getLastRow();
  if (lastRowConfig < 2) return;
  const configs = configSheet.getRange(2, 1, lastRowConfig - 1, 3).getValues()
    .filter(c => c[0] && c[2]);

  const now = new Date();
  const reportData = {};
  let hasNewArticles = false;
  const yesterdayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 0, 0, 0).getTime();
  const yesterdayEnd   = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 23, 59, 59, 999).getTime();
  const allNewArticlesForGCS = [];

  const isLocalArticle = (title) => {
    const text = String(title);
    return LOCATION_PLUS_WORDS.some(w => text.includes(w));
  };

  const calcScore = (title, category) => {
    let score = 0;
    const text = String(title);
    (PLUS_WORDS[category] || []).forEach(w => { if (text.includes(w)) score += 1; });
    STRONG_WORDS.forEach(w => { if (text.includes(w)) score += 2; });
    EXCLUDE_WORDS.forEach(w => { if (text.includes(w)) score -= 3; });
    if (!isLocalArticle(text)) {
      NATIONAL_BRAND_WORDS.forEach(w => { if (text.includes(w)) score += 1; });
    }
    return score;
  };

  const cleanText = (t) => String(t).replace(/[^a-zA-Z0-9ａ-ｚＡ-Ｚ０-９ぁ-んァ-ヶ一-龠]/g, '').substring(0, 10);

  const requests = configs.map(c => ({ url: c[2], method: 'get', muteHttpExceptions: true }));
  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    console.error('RSS fetchAll失敗: ' + e.toString());
    return;
  }

  configs.forEach((config, idx) => {
    const category = config[0];
    const res = responses[idx];

    try {
      if (!res || res.getResponseCode() !== 200) {
        console.warn(`${category}: HTTP ${res ? res.getResponseCode() : 'null'}`);
        return;
      }

      let targetSheet = ss.getSheetByName(category);
      if (!targetSheet) {
        targetSheet = ss.insertSheet(category);
        targetSheet.getRange(1, 1, 1, 5).setValues([["取得日時", "カテゴリ", "記事タイトル", "URL", "ID"]]);
      }
      const lastRowTarget = targetSheet.getLastRow();

      const rawLimit = CATEGORY_LIMITS[category];
      const limits = typeof rawLimit === 'number'
        ? { local: Math.ceil(rawLimit / 2), national: Math.floor(rawLimit / 2) }
        : (rawLimit || { local: 3, national: 2 });
      const totalLimit = limits.local + limits.national;

      const existingIds = lastRowTarget >= 1 ? targetSheet.getRange(1, 5, lastRowTarget, 1).getValues().flat() : [];
      const idSet = new Set(existingIds);
      const existingTitles = lastRowTarget >= 1 ? targetSheet.getRange(1, 3, lastRowTarget, 1).getValues().flat() : [];
      const titlePrefixSet = new Set(existingTitles.map(cleanText));

      const xml = XmlService.parse(res.getContentText());
      const items = xml.getRootElement().getChild('channel').getChildren('item');

      const pending = [];
      items.forEach(item => {
        const title = item.getChildText('title');
        const link = item.getChildText('link');
        const guid = item.getChildText('guid');
        const pubDateStr = item.getChildText('pubDate');
        const pubDate = pubDateStr ? new Date(pubDateStr).getTime() : 0;
        if (pubDate < yesterdayStart || pubDate > yesterdayEnd) return;
        const prefix = cleanText(title);
        if (idSet.has(guid) || titlePrefixSet.has(prefix)) return;
        const score = calcScore(title, category);
        if (score >= 0) {
          pending.push({
            title, link, guid, prefix, score, pubDate,
            isLocal: isLocalArticle(title)
          });
        }
      });

      const sortFn = (a, b) =>
        b.score !== a.score ? b.score - a.score : b.pubDate - a.pubDate;
      const localPool    = pending.filter(a => a.isLocal).sort(sortFn);
      const nationalPool = pending.filter(a => !a.isLocal).sort(sortFn);

      let selected = [
        ...localPool.slice(0, limits.local),
        ...nationalPool.slice(0, limits.national)
      ];

      if (selected.length < totalLimit) {
        const overflow = [
          ...localPool.slice(limits.local),
          ...nationalPool.slice(limits.national)
        ].sort(sortFn);
        selected = selected.concat(overflow.slice(0, totalLimit - selected.length));
      }

      if (selected.length === 0) return;

      selected.sort(sortFn);

      const newRows = [];
      const articles = [];
      selected.forEach(art => {
        newRows.push([now, category, art.title, art.link, art.guid]);
        articles.push({ title: art.title, link: art.link });
        allNewArticlesForGCS.push({
          fetched_at: now, category, title: art.title,
          url: art.link, guid: art.guid, scope: art.isLocal ? 'local' : 'national'
        });
      });
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
      reportData[category] = articles;
      hasNewArticles = true;

    } catch (e) {
      console.error(`エラー発生(${category}): ${e.toString()}`);
    }
  });

  if (allNewArticlesForGCS.length > 0) {
    const timestamp = Utilities.formatDate(now, "JST", "yyyyMMdd_HHmmss");
    uploadToGCS(`backup/news_${timestamp}.json`, allNewArticlesForGCS);
  }

  if (hasNewArticles) {
    sendEmailReport(reportData, now, ss);
  } else {
    console.log('新着記事なし — メール配信スキップ');
  }
}

// ──────────────────────────────────────────────────────────────
//  GCSバックアップ
//  BUCKET_NAME が未設定の場合は何もせず終了
// ──────────────────────────────────────────────────────────────
function uploadToGCS(fileName, data) {
  const bucketName = PropertiesService.getScriptProperties().getProperty('BUCKET_NAME');
  if (!bucketName) {
    console.log('BUCKET_NAME未設定のためGCSバックアップをスキップ');
    return;
  }
  const url = `https://storage.googleapis.com/upload/storage/v1/b/${bucketName}/o?uploadType=media&name=${fileName}`;
  const token = ScriptApp.getOAuthToken();
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    console.error(`GCS保存失敗: ${response.getContentText()}`);
  }
}

// ──────────────────────────────────────────────────────────────
//  HTMLメール配信
// ──────────────────────────────────────────────────────────────
function sendEmailReport(reportData, now, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  reportData = reportData || {};
  now = now || new Date();

  const emailSheet = ss.getSheetByName('配信リスト');
  if (!emailSheet) return;
  const lastRowEmail = emailSheet.getLastRow();
  if (lastRowEmail < 2) return;
  const emailList = emailSheet.getRange(2, 2, lastRowEmail - 1, 1)
    .getValues().flat().filter(e => e !== '');
  if (emailList.length === 0) return;

  const CATEGORY_STYLE = {
    'お得情報・クーポン':       { color: '#9d174d', border: '#ec4899', dot: '#ec4899' },
    'ハッカソン・コンテスト':   { color: '#1d4ed8', border: '#3b82f6', dot: '#3b82f6' },
    '留学・海外プログラム':     { color: '#047857', border: '#10b981', dot: '#10b981' },
    'インターン・プログラム':   { color: '#c2410c', border: '#f97316', dot: '#f97316' },
  };
  const DEFAULT_STYLE = { color: '#374151', border: '#9ca3af', dot: '#9ca3af' };

  const CATEGORY_ORDER = [
    'お得情報・クーポン',
    'ハッカソン・コンテスト',
    '留学・海外プログラム',
    'インターン・プログラム'
  ];

  let categorySections = '';
  const orderedEntries = CATEGORY_ORDER
    .filter(c => reportData[c] && reportData[c].length > 0)
    .map(c => [c, reportData[c]]);
  for (const c of Object.keys(reportData)) {
    if (!CATEGORY_ORDER.includes(c) && reportData[c] && reportData[c].length > 0) {
      orderedEntries.push([c, reportData[c]]);
    }
  }

  for (const [category, articles] of orderedEntries) {
    if (!Array.isArray(articles) || articles.length === 0) continue;
    const s = CATEGORY_STYLE[category] || DEFAULT_STYLE;

    const articleRows = articles.map(article => {
      if (!article) return '';
      let source = '';
      try {
        const host = new URL(article.link).hostname.replace('www.', '');
        source = host.includes('google.com') ? '' : host;
      } catch(e) {}
      const url = article.link || '#';
      return `
        <tr>
          <td width="12" valign="top" style="padding:8px 0;">
            <div style="width:4px;height:4px;border-radius:50%;background:${s.dot};
                        margin-top:7px;opacity:.5;"></div>
          </td>
          <td style="padding:8px 0 8px 8px;border-bottom:1px solid #f9fafb;">
            <a href="${url}" target="_blank"
               style="display:block;text-decoration:none;color:inherit;
                      font-family:'Hiragino Kaku Gothic ProN','Noto Sans JP',sans-serif;">
              <span style="font-size:13px;color:#111827;line-height:1.6;
                           display:block;margin-bottom:2px;">${article.title || ''}</span>
              ${source ? `<span style="font-size:11px;color:#9ca3af;">${source}</span>` : ''}
            </a>
          </td>
        </tr>`;
    }).join('');

    categorySections += `
      <tr><td style="padding:0 0 20px 0;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td style="padding-bottom:10px;border-bottom:2px solid ${s.border};">
              <table cellpadding="0" cellspacing="0"><tr>
                <td><div style="width:7px;height:7px;border-radius:50%;background:${s.dot};"></div></td>
                <td style="padding-left:8px;">
                  <span style="font-size:11px;font-weight:700;letter-spacing:.08em;
                               text-transform:uppercase;color:${s.color};">${category}</span>
                </td>
                <td style="padding-left:12px;">
                  <span style="font-size:10px;color:#9ca3af;">${articles.length}件</span>
                </td>
              </tr></table>
            </td>
          </tr>
          <tr><td>
            <table width="100%" cellpadding="0" cellspacing="0">${articleRows}</table>
          </td></tr>
        </table>
      </td></tr>`;
  }

  const formattedDate     = Utilities.formatDate(now, 'JST', 'yyyy年MM月dd日（E）');
  const formattedDatetime = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm');

  const htmlBody = `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background:#f9fafb;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f9fafb;">
<tr><td align="center" style="padding:24px 16px;">
<table width="600" cellpadding="0" cellspacing="0" border="0"
       style="max-width:600px;width:100%;background:#ffffff;
              border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">
 
  <tr><td style="background:#111827;padding:24px 28px 20px;">
    <p style="font-size:20px;font-weight:700;color:#f9fafb;margin:0 0 8px 0;
              font-family:'Hiragino Kaku Gothic ProN','Noto Sans JP',sans-serif;">福岡学生情報マガジン</p>
    <p style="font-size:12px;color:#6b7280;margin:0;font-family:sans-serif;">
      ${formattedDate}</p>
  </td></tr>
 
  <tr><td style="padding:24px 28px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      ${categorySections}
    </table>
  </td></tr>
 
  <tr><td style="background:#f9fafb;border-top:1px solid #e5e7eb;padding:14px 28px;" align="right">
    <span style="font-size:11px;color:#d1d5db;font-family:sans-serif;">${formattedDatetime}</span>
  </td></tr>
 
</table>
</td></tr>
</table>
</body>
</html>`;

  const subject = `【福岡学生情報】お得・チャンス・学びまとめ ${Utilities.formatDate(now, 'JST', 'MM/dd')}`;
  try {
    GmailApp.sendEmail(emailList.join(','), subject, '', { htmlBody: htmlBody });
  } catch (e) {
    console.error('メール送信エラー: ' + e.toString());
  }
}
