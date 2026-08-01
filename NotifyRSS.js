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

      let iconUrl = customIconUrl;
      if (!iconUrl) {
        const domain = getDomainFromUrl(rssUrl);
        iconUrl = domain ? `https://icons.duckduckgo.com/ip3/${domain}.ico` : '';
      }

      let latestPublishedTime = lastNotifiedTime;
      const newPosts = [];

      items.forEach(item => {
        const title = getElementTextByNames(item, ['title']);
        const link = getLinkFromItem(item);
        const pubDateStr = getElementTextByNames(item, ['date', 'pubdate', 'published', 'updated']);
        const imageUrl = getImageUrlFromItem(item); // 記事画像（アイキャッチ）の取得
        
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
        console.log(`[新着あり] 行 ${rowIndex}: "${siteName}" から ${newPosts.length} 件の新規記事を送信します...`);
        
        newPosts.sort((a, b) => a.pubTime - b.pubTime);

        newPosts.forEach((post, postIndex) => {
          console.log(`  └ [送信 ${postIndex + 1}/${newPosts.length}] "${post.title}"`);
          sendDiscordEmbedNotification(webhookUrl, post.title, post.link, post.pubTime, siteName, iconUrl, colorHex, post.imageUrl);
          Utilities.sleep(1000);
        });

        // H列の最終取得日時を更新
        sheet.getRange(rowIndex, 8).setValue(new Date(latestPublishedTime));
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
function sendDiscordEmbedNotification(webhookUrl, title, link, pubTime, siteName, iconUrl, hexColorStr, imageUrl) {
  const safeLink = encodeURI(link);

  // 16進数カラーコードを整数（DEC）に変換
  let colorNum = 0x3498db; // デフォルト：青
  if (hexColorStr) {
    const cleanHex = String(hexColorStr).replace('#', '').trim();
    // 長さが6桁かつ16進数として正しく変換できるか判定（正規表現不使用）
    if (cleanHex.length === 6 && !isNaN(parseInt(cleanHex, 16))) {
      colorNum = parseInt(cleanHex, 16);
    }
  }

  // 投稿時刻（ISO 8601フォーマット）
  const isoTimestamp = new Date(pubTime > 0 ? pubTime : Date.now()).toISOString();

  // Embed構造の作成
  const embed = {
    title: title,
    url: safeLink,
    color: colorNum,
    footer: {
      text: siteName || "RSS Feed",
      icon_url: iconUrl || undefined
    },
    timestamp: isoTimestamp
  };

  // サムネイル画像（アイキャッチ）があれば追加
  if (imageUrl) {
    embed.thumbnail = { url: imageUrl };
  }

  const payload = {
    content: title, // カードの上にタイトルテキストを配置
    embeds: [embed]  // 自前カード
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch(webhookUrl, options);
}

/**
 * RSSアイテムから画像を抽出する関数
 */
function getImageUrlFromItem(item) {
  const children = item.getChildren();
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    const name = child.getName().toLowerCase();

    // <media:thumbnail url="..." /> または <media:content url="..." />
    if (name === 'thumbnail' || name === 'content') {
      const urlAttr = child.getAttribute('url');
      if (urlAttr) return urlAttr.getValue();
    }
    // <enclosure url="..." type="image/..." />
    if (name === 'enclosure') {
      const typeAttr = child.getAttribute('type');
      const urlAttr = child.getAttribute('url');
      if (urlAttr && typeAttr && typeAttr.getValue().startsWith('image/')) {
        return urlAttr.getValue();
      }
    }
  }

  // 本文（description）内の <img> タグからsrcを抽出
  const description = getElementTextByNames(item, ['description', 'content']);
  if (description) {
    const imgMatch = description.match(/<img[^>]+src=["']([^"']+)["']/i);
    if (imgMatch && imgMatch[1]) {
      return imgMatch[1];
    }
  }

  return null;
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