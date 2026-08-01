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

  // D列(4)～H列(8)の5列分を取得
  // D: RSS URL, E: Webhook URL, F: Icon URL, G: Color, H: 最終更新日時
  const dataRange = sheet.getRange(2, 4, lastRow - 1, 5);
  const values = dataRange.getValues();

  values.forEach((row, index) => {
    const rssUrl = row[0];
    const webhookUrl = row[1];
    const customIconUrl = row[2]; // F列: カスタムアイコンURL
    const colorHex = row[3];      // G列: 16進数カラーコード (#FF0000等)
    const rawLastNotified = row[4]; // H列: 最終更新日時
    
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
        
        // 強化された画像抽出ロジック
        const imageUrl = getImageUrlFromItem(item, title);
        
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
          newPosts.push({ title, link, pubTime, imageUrl });
          if (pubTime > latestPublishedTime) {
            latestPublishedTime = pubTime;
          }
        }
      });

      // 初回登録処理（H列に書き込み）
      if (isFirstTime) {
        const recordTime = latestPublishedTime > 0 ? new Date(latestPublishedTime) : new Date();
        sheet.getRange(rowIndex, 8).setValue(recordTime); // H列(8)に更新日時を保存
        console.log(`[初回登録] 行 ${rowIndex}: "${siteName}" の基準日時を記録しました。基準日時: ${recordTime}`);
        return;
      }

      // 2回目以降：通知送信
      if (newPosts.length > 0) {
        console.log(`[新着あり] 行 ${rowIndex}: "${siteName}" から ${newPosts.length} 件の新規記事があります。`);
        
        // 古い順にソート
        newPosts.sort((a, b) => a.pubTime - b.pubTime);

        // レートリミット回避：一度に送信するのは最大5件まで
        const postsToSend = newPosts.slice(0, 5);
        if (newPosts.length > 5) {
          console.warn(`  └ ※件数が多いため、今回の実行では先頭5件のみ送信します（残りは次回）。`);
        }

        postsToSend.forEach((post, postIndex) => {
          console.log(`  └ [送信 ${postIndex + 1}/${postsToSend.length}] "${post.title}" | 画像: ${post.imageUrl ? '〇' : '×'}`);
          sendDiscordEmbedNotification(webhookUrl, post.title, post.link, post.pubTime, siteName, customIconUrl, defaultIconUrl, colorHex, post.imageUrl);
          
          // レートリミット（429エラー）回避のため 2.5秒 待機
          Utilities.sleep(2500);
        });

        // H列の最終取得日時を更新（送信した最後の記事の日時を記録）
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
 * Discord WebhookへEmbed（カード）形式で通知を送信する関数
 */
function sendDiscordEmbedNotification(webhookUrl, title, link, pubTime, siteName, customIconUrl, defaultIconUrl, hexColorStr, imageUrl) {
  const safeLink = encodeURI(link);

  let colorNum = 0x3498db; // デフォルト：青
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

  // メイン画像（カード下部の大きな画像）として追加
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
  const responseCode = response.getResponseCode();

  if (responseCode < 200 || responseCode >= 300) {
    console.error(`[Discord送信失敗] HTTP ${responseCode}: ${response.getContentText()}`);
  }
}

/**
 * RSSアイテムから画像を抽出する関数（ライブドアブログ・各種メディア強化版）
 */
function getImageUrlFromItem(item, itemTitle) {
  const children = item.getChildren();

  // 1. 専用の画像タグから抽出 (<media:thumbnail>, <media:content>, <enclosure>, <image>)
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    const name = child.getName().toLowerCase();

    // <media:thumbnail url="..."> / <media:content url="...">
    if (name === 'thumbnail' || name === 'content') {
      const urlAttr = child.getAttribute('url');
      if (urlAttr && urlAttr.getValue()) {
        console.log(`[画像抽出] <media:${name}> から取得 (${itemTitle}): ${urlAttr.getValue()}`);
        return urlAttr.getValue();
      }
    }
    // <enclosure url="..." type="image/...">
    if (name === 'enclosure') {
      const typeAttr = child.getAttribute('type');
      const urlAttr = child.getAttribute('url');
      if (urlAttr && urlAttr.getValue()) {
        if (!typeAttr || typeAttr.getValue().startsWith('image/')) {
          console.log(`[画像抽出] <enclosure> から取得 (${itemTitle}): ${urlAttr.getValue()}`);
          return urlAttr.getValue();
        }
      }
    }
    // <image> タグの直下に URL テキストがある場合（一部のRDF等）
    if (name === 'image') {
      const imgText = child.getText();
      if (imgText && imgText.startsWith('http')) {
        console.log(`[画像抽出] <image> から取得 (${itemTitle}): ${imgText}`);
        return imgText;
      }
    }
  }

  // 2. 本文HTMLから <img> タグを抽出する（encoded -> content -> description の優先順で探索）
  const htmlContents = getAllHtmlContents(item);

  for (let j = 0; j < htmlContents.length; j++) {
    let rawHtml = htmlContents[j];
    if (!rawHtml) continue;

    // HTML特殊文字のエスケープ解除
    rawHtml = rawHtml
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#039;/g, "'")
      .replace(/&amp;/g, '&');

    // <img> タグから src 属性を取り出す正規表現（複数マッチ）
    const imgRegex = /<img[^>]+src=["']([^"']+)["']/gi;
    let match;

    while ((match = imgRegex.exec(rawHtml)) !== null) {
      const imgSrc = match[1];

      // ブログのカウンター用画像、絵文字、アクセス解析用ドット画像を除外
      if (
        imgSrc &&
        !imgSrc.includes('/smilies/') &&
        !imgSrc.includes('emoji') &&
        !imgSrc.includes('counter') &&
        !imgSrc.includes('analy') &&
        !imgSrc.includes('.gif') // アイコン系のGIF画像を除外
      ) {
        console.log(`[画像抽出] 本文HTMLから取得 (${itemTitle}): ${imgSrc}`);
        return imgSrc;
      }
    }
  }

  console.log(`[画像抽出なし] 画像が見つかりませんでした (${itemTitle})`);
  return null;
}

/**
 * item配下から encoded / content / description などのHTML全文を配列で取得
 */
function getAllHtmlContents(parentElement) {
  const children = parentElement.getChildren();
  const contents = [];

  // 優先順位: encoded（全文） -> content -> description（概要）
  const targetNames = ['encoded', 'content', 'description'];

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

function getLinkFromItem(item) {
  const linkText = getElementTextByNames(item, ['link']);
  if (linkText) return linkText;

  const children = item.getChildren();
  for (let i = 0; i < children.length; i++) {
    if (children[i].getName().toLowerCase() === 'link') {
      const href = children[i].getAttribute('href');
      if (href) return href.getValue();
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