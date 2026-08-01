/**
 * スプレッドシートからRSSとDiscord Webhookを読み込み、Discord Embedカードで通知する関数
 */
function checkRssAndNotifyDiscord() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    console.error("エラー: スクリプトプロパティ 'SPREADSHEET_ID' が設定されていません。");
    return;
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return;

  const dataRange = sheet.getRange(2, 4, lastRow - 1, 5);
  const values = dataRange.getValues();

  values.forEach((row, index) => {
    const rssUrl = row[0];
    const webhookUrl = row[1];
    const customIconUrl = row[2];
    const colorHex = row[3];
    const rawLastNotified = row[4];
    
    const isFirstTime = !rawLastNotified;
    const lastNotifiedTime = rawLastNotified ? new Date(rawLastNotified).getTime() : 0;
    const rowIndex = index + 2;

    if (!rssUrl || !webhookUrl) return;

    try {
      const fetchOptions = {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
      };

      const response = UrlFetchApp.fetch(rssUrl, fetchOptions);
      if (response.getResponseCode() !== 200) {
        console.warn(`[行 ${rowIndex}] HTTPエラー ${response.getResponseCode()} (${rssUrl})`);
        return;
      }

      const xml = XmlService.parse(response.getContentText());
      const root = xml.getRootElement();
      const rootName = root.getName().toLowerCase();
      
      let items = [];
      let siteName = '';

      if (rootName === 'rss') {
        const channel = getChildByLocalName(root, 'channel');
        if (channel) {
          items = getChildrenByLocalName(channel, 'item');
          siteName = getElementTextByNames(channel, ['title']);
        }
      } else if (rootName === 'rdf') {
        items = getChildrenByLocalName(root, 'item');
        const channel = getChildByLocalName(root, 'channel');
        if (channel) {
          siteName = getElementTextByNames(channel, ['title']);
        }
      } else if (rootName === 'feed') {
        items = getChildrenByLocalName(root, 'entry');
        siteName = getElementTextByNames(root, ['title']);
      }

      const domain = getDomainFromUrl(rssUrl);
      const defaultIconUrl = domain ? `https://icons.duckduckgo.com/ip3/${domain}.ico` : '';

      let latestPublishedTime = lastNotifiedTime;
      const newPosts = [];

      items.forEach(item => {
        const title = getElementTextByNames(item, ['title']);
        const link = getLinkFromItem(item);
        const pubDateStr = getElementTextByNames(item, ['date', 'pubdate', 'published', 'updated']);
        
        let pubTime = 0;
        if (pubDateStr) {
          pubTime = new Date(pubDateStr).getTime();
        }
        
        if (isNaN(pubTime) || pubTime === 0) {
          pubTime = lastNotifiedTime + 1;
        }

        if (isFirstTime) {
          if (pubTime > latestPublishedTime) {
            latestPublishedTime = pubTime;
          }
        } else if (pubTime > lastNotifiedTime) {
          // 新規記事の場合のみ画像抽出を実施（効率化のため）
          const imageUrl = getImageUrlFromItemOrWeb(item, title, link);
          newPosts.push({ title, link, pubTime, imageUrl });
          if (pubTime > latestPublishedTime) {
            latestPublishedTime = pubTime;
          }
        }
      });

      // 初回登録処理
      if (isFirstTime) {
        const recordTime = latestPublishedTime > 0 ? new Date(latestPublishedTime) : new Date();
        sheet.getRange(rowIndex, 8).setValue(recordTime);
        console.log(`[初回登録] 行 ${rowIndex}: "${siteName}" の基準日時を記録しました。基準日時: ${recordTime}`);
        return;
      }

      // 2回目以降：通知送信
      if (newPosts.length > 0) {
        console.log(`[新着あり] 行 ${rowIndex}: "${siteName}" から ${newPosts.length} 件の新規記事があります。`);
        
        newPosts.sort((a, b) => a.pubTime - b.pubTime);

        const postsToSend = newPosts.slice(0, 5);
        if (newPosts.length > 5) {
          console.warn(`  └ ※件数が多いため、今回の実行では先頭5件のみ送信します。`);
        }

        postsToSend.forEach((post, postIndex) => {
          console.log(`  └ [送信 ${postIndex + 1}/${postsToSend.length}] "${post.title}" | 画像: ${post.imageUrl ? post.imageUrl : 'なし'}`);
          sendDiscordEmbedNotification(webhookUrl, post.title, post.link, post.pubTime, siteName, customIconUrl, defaultIconUrl, colorHex, post.imageUrl);
          
          Utilities.sleep(2500);
        });

        const lastSentPost = postsToSend[postsToSend.length - 1];
        sheet.getRange(rowIndex, 8).setValue(new Date(lastSentPost.pubTime));
        console.log(`  └ 完了: H列の最終更新日時を更新しました。`);
      }

    } catch (error) {
      console.error(`[エラー] 行 ${rowIndex} の処理中に例外が発生しました (${rssUrl}):`, error);
    }
  });
}

/**
 * RSSおよびWebページから画像を最適抽出するハイブリッド関数
 */
function getImageUrlFromItemOrWeb(item, itemTitle, itemLink) {
  // 1. XMLタグ構造からの判定
  const foundInXml = findImageInElementRecursive(item);
  if (foundInXml) {
    console.log(`[画像抽出:XML構造] (${itemTitle}): ${foundInXml}`);
    return foundInXml;
  }

  // 2. 本文HTMLからの判定（PochippやAmazonを除外）
  const htmlContents = getAllHtmlContents(item);
  for (let j = 0; j < htmlContents.length; j++) {
    let rawHtml = decodeHtmlEntitiesFully(htmlContents[j]);
    if (!rawHtml) continue;

    const imgRegex = /<img[^>]+(?:src|data-src|data-original|srcset)=["']([^"'\s>]+)["']/gi;
    let match;

    while ((match = imgRegex.exec(rawHtml)) !== null) {
      let imgSrc = match[1];

      if (imgSrc.includes(',')) imgSrc = imgSrc.split(',')[0].trim().split(' ')[0];
      if (imgSrc.startsWith('//')) imgSrc = 'https:' + imgSrc;

      if (isValidImageUrl(imgSrc)) {
        console.log(`[画像抽出:RSS本文] (${itemTitle}): ${imgSrc}`);
        return imgSrc;
      }
    }
  }

  // 3. RSS内に画像がない場合、記事URL（Webページ）から og:image を取得
  if (itemLink && (itemLink.startsWith('http://') || itemLink.startsWith('https://'))) {
    const ogImage = fetchOgImageFromUrl(itemLink);
    if (ogImage) {
      console.log(`[画像抽出:WebページOGP] (${itemTitle}): ${ogImage}`);
      return ogImage;
    }
  }

  console.log(`[画像抽出なし] (${itemTitle})`);
  return null;
}

/**
 * 記事ページにアクセスして og:image を取得する関数
 */
function fetchOgImageFromUrl(url) {
  try {
    const options = {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
      }
    };
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() !== 200) return null;

    const html = res.getContentText();
    // og:image タグを抽出
    const ogMatch = html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i) ||
                    html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image["']/i);

    if (ogMatch && ogMatch[1]) {
      let ogUrl = ogMatch[1].trim();
      if (ogUrl.startsWith('//')) ogUrl = 'https:' + ogUrl;
      if (isValidImageUrl(ogUrl)) return ogUrl;
    }
  } catch (e) {
    console.warn(`[OGP取得失敗] ${url}: ${e.message}`);
  }
  return null;
}

/**
 * 有効なアイキャッチ画像か判定（ポチップ、Amazon、アイコン、広告等を厳格に遮断）
 */
function isValidImageUrl(url) {
  if (!url || typeof url !== 'string' || (!url.startsWith('http://') && !url.startsWith('https://'))) {
    return false;
  }

  const lower = url.toLowerCase();

  // 強力除外キーワード（ポチップロゴ、Amazon、プラグイン、アフィリエイト、SNSアイコン等）
  const ignoreKeywords = [
    'amazon.com', 'amazon-adsystem.com', 'm.media-amazon.com', 'ssl-images-amazon.com',
    'pochipp', 'pochipp-logo', 'plugins/pochipp', // ポチッププラグイン画像
    'wp-includes', '/smilies/', 'emoji', 'counter', 'analy', 'facebook.com', 'twitter.com', 
    'line.me', 'hatena', 'share', 'avatar', 'button', 'icon', 'clear.gif', 
    'blank.gif', 'pixel', 'ad_banner', 'banner', 'widgets', 'logo_publisher'
  ];

  for (let i = 0; i < ignoreKeywords.length; i++) {
    if (lower.includes(ignoreKeywords[i])) return false;
  }

  if (lower.split('?')[0].endsWith('.gif')) return false;

  return (
    lower.includes('.jpg') ||
    lower.includes('.jpeg') ||
    lower.includes('.png') ||
    lower.includes('.webp') ||
    lower.includes('wp-content/uploads')
  );
}

/**
 * XML要素および属性を再帰探索
 */
function findImageInElementRecursive(element) {
  const name = element.getName().toLowerCase();

  const attributes = element.getAttributes();
  for (let i = 0; i < attributes.length; i++) {
    const val = attributes[i].getValue();
    if (val) {
      let cleanVal = val.replace(/[\r\n\t]/g, '').trim();
      if (cleanVal.startsWith('//')) cleanVal = 'https:' + cleanVal;
      
      if (['thumbnail', 'content', 'enclosure', 'image'].includes(name) || attributes[i].getName().toLowerCase() === 'url') {
        if (isValidImageUrl(cleanVal)) {
          return cleanVal;
        }
      }
    }
  }

  if (['image', 'thumbnail', 'enclosure'].includes(name)) {
    const text = element.getText();
    if (text) {
      let cleanText = text.replace(/[\r\n\t]/g, '').trim();
      if (cleanText.startsWith('//')) cleanText = 'https:' + cleanText;
      if (isValidImageUrl(cleanText)) {
        return cleanText;
      }
    }
  }

  const children = element.getChildren();
  for (let i = 0; i < children.length; i++) {
    const res = findImageInElementRecursive(children[i]);
    if (res) return res;
  }

  return null;
}

/**
 * HTMLテキストのマルチデコード
 */
function decodeHtmlEntitiesFully(str) {
  if (!str) return '';
  let decoded = str;
  for (let i = 0; i < 3; i++) {
    decoded = decoded
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#039;/g, "'")
      .replace(/&#39;/g, "'")
      .replace(/&amp;/g, '&');
  }
  return decoded;
}

/**
 * item配下から本文テキスト領域を全取得
 */
function getAllHtmlContents(parentElement) {
  const children = parentElement.getChildren();
  const contents = [];
  const targetNames = ['encoded', 'content', 'description', 'summary'];

  targetNames.forEach(target => {
    for (let i = 0; i < children.length; i++) {
      if (children[i].getName().toLowerCase() === target) {
        const text = children[i].getText();
        if (text && text.trim() !== '') {
          contents.push(text);
        }
      }
    }
  });

  return contents;
}

/**
 * リンクURLを取得
 */
function getLinkFromItem(item) {
  let rawLink = '';
  const linkText = getElementTextByNames(item, ['link']);
  if (linkText) {
    rawLink = linkText;
  } else {
    const children = item.getChildren();
    for (let i = 0; i < children.length; i++) {
      if (children[i].getName().toLowerCase() === 'link') {
        const hrefAttr = children[i].getAttribute('href');
        if (hrefAttr && hrefAttr.getValue()) {
          rawLink = hrefAttr.getValue();
          break;
        }
      }
    }
  }

  return rawLink ? rawLink.replace(/[\r\n\t]/g, '').trim() : '';
}

/**
 * Discord通知機能
 */
function sendDiscordEmbedNotification(webhookUrl, title, link, pubTime, siteName, customIconUrl, defaultIconUrl, hexColorStr, imageUrl) {
  const safeLink = link.replace(/[\r\n\t]/g, '').trim();

  let colorNum = 0x3498db;
  if (hexColorStr) {
    const cleanHex = String(hexColorStr).replace('#', '').trim();
    if (cleanHex.length === 6 && !isNaN(parseInt(cleanHex, 16))) {
      colorNum = parseInt(cleanHex, 16);
    }
  }

  const isoTimestamp = new Date(pubTime > 0 ? pubTime : Date.now()).toISOString();
  const footerIcon = customIconUrl || defaultIconUrl;

  const embed = {
    title: title,
    url: safeLink,
    color: colorNum,
    footer: {
      text: siteName || "RSS Feed",
      icon_url: footerIcon || undefined
    },
    timestamp: isoTimestamp
  };

  if (imageUrl) {
    embed.image = { url: imageUrl };
  }

  const payload = {
    content: title,
    embeds: [embed]
  };

  if (siteName && siteName.trim() !== '') {
    payload.username = siteName.substring(0, 80);
  }
  
  if (customIconUrl && (customIconUrl.startsWith('http://') || customIconUrl.startsWith('https://'))) {
    payload.avatar_url = customIconUrl;
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    console.error(`[Discord送信失敗] HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
  }
}

/* ヘルパー関数群 */
function getChildByLocalName(parentElement, localName) {
  const children = parentElement.getChildren();
  for (let i = 0; i < children.length; i++) {
    if (children[i].getName().toLowerCase() === localName.toLowerCase()) return children[i];
  }
  return null;
}

function getChildrenByLocalName(parentElement, localName) {
  const children = parentElement.getChildren();
  const result = [];
  for (let i = 0; i < children.length; i++) {
    if (children[i].getName().toLowerCase() === localName.toLowerCase()) result.push(children[i]);
  }
  return result;
}

function getElementTextByNames(parentElement, tagNames) {
  const children = parentElement.getChildren();
  for (let i = 0; i < children.length; i++) {
    const childName = children[i].getName().toLowerCase();
    if (tagNames.some(name => name.toLowerCase() === childName)) {
      return children[i].getText();
    }
  }
  return '';
}

function getDomainFromUrl(url) {
  try {
    const matches = url.match(/^https?:\/\/([^/?#]+)(?:[/?#]|$)/i);
    return matches && matches[1];
  } catch (e) {
    return '';
  }
}

/**
 * 抽出できないサイトのRSS構造を調査するためのデバッグ用関数
 */
function debugRssFeeds() {
  // 調査したいサイトのRSS URL一覧
  const testUrls = [
    'https://mega-nikke.com/feed',             // メガニケ速報
    'https://gigazine.net/news/rss_2.0/',      // GIGAZINE
    'https://www.4gamer.net/rss/index.xml',     // 4Gamer
    'https://game.watch.impress.co.jp/data/rss/1.0/gw/feed.rdf', // Game Watch
    'https://daily-guraburu.com/feed'       // 日々是グラブる
  ];

  testUrls.forEach(url => {
    console.log(`========================================`);
    console.log(`[調査開始] URL: ${url}`);
    console.log(`========================================`);

    try {
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
      });

      const xmlText = response.getContentText();
      
      // 先頭の1件（<item> または <entry>）のテキストを抽出してログ表示
      const itemMatch = xmlText.match(/<(item|entry)[\s\S]*?<\/(item|entry)>/i);
      if (itemMatch) {
        console.log(`--- 先頭記事の生XML (最初の1000文字) ---`);
        console.log(itemMatch[0].substring(0, 1000));

        // 含まれている img src や url 属性をすべて列挙
        console.log(`--- 見つかった画像URL一覧 ---`);
        const imgMatches = itemMatch[0].match(/https?:\/\/[^\s"'<>]+\.(jpg|jpeg|png|webp|gif)/gi);
        if (imgMatches) {
          const uniqueImgs = [...new Set(imgMatches)];
          uniqueImgs.forEach(img => console.log(` - ${img}`));
        } else {
          console.log(" ※画像拡張子を含むURLが見つかりませんでした。");
        }
      } else {
        console.log("記事（item/entry）タグが見つかりませんでした。");
      }

    } catch (e) {
      console.error(`取得失敗: ${e.message}`);
    }
  });
}