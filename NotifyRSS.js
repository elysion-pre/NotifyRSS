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
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'application/rss+xml, application/xml, text/xml, */*'
        }
      };

      const response = UrlFetchApp.fetch(rssUrl, fetchOptions);
      if (response.getResponseCode() !== 200) {
        console.warn(`[行 ${rowIndex}] HTTPエラー ${response.getResponseCode()} (${rssUrl})`);
        return;
      }

      // XMLパースの安全性強化（構文エラーのあるRSS対策）
      let xml;
      try {
        xml = XmlService.parse(response.getContentText());
      } catch (xmlError) {
        console.error(`[行 ${rowIndex}] XMLのパースに失敗しました (${rssUrl}): ${xmlError.message}`);
        return;
      }

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
          console.log(`  └ [送信 ${postIndex + 1}/${postsToSend.length}] "${post.title}"`);
          sendDiscordEmbedNotification(webhookUrl, post.title, post.link, post.pubTime, siteName, customIconUrl, defaultIconUrl, colorHex, post.imageUrl);
          
          Utilities.sleep(2000); // Discordのレートリミット回避用のウェイト
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
 * 画像抽出ロジック（WebページOGP優先＆RSS解析フォールバック）
 */
function getImageUrlFromItemOrWeb(item, itemTitle, itemLink) {
  // 優先度1: 記事ページへ直接アクセスして og:image を取得
  if (itemLink && (itemLink.startsWith('http://') || itemLink.startsWith('https://'))) {
    const ogImage = fetchOgImageFromUrl(itemLink);
    if (ogImage) {
      console.log(`[画像抽出:WebページOGP] (${itemTitle}): ${ogImage}`);
      return ogImage;
    }
  }

  // 優先度2: RSS内のアイキャッチメタデータ(media:content, media:thumbnail, enclosure)を探す
  const metaImage = findFeaturedImageInXml(item);
  if (metaImage) {
    console.log(`[画像抽出:アイキャッチメタデータ] (${itemTitle}): ${metaImage}`);
    return metaImage;
  }

  // 優先度3: RSS本文(encoded/description) 内から画像を順次抽出
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

  console.log(`[画像抽出なし] (${itemTitle})`);
  return null;
}

/**
 * XMLからアイキャッチ画像メタデータ(media:content/thumbnail, enclosure)を検索
 */
function findFeaturedImageInXml(element) {
  const children = element.getChildren();

  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    const name = child.getName().toLowerCase();
    
    if (name === 'content' || name === 'thumbnail') {
      const urlAttr = child.getAttribute('url');
      if (urlAttr && isValidImageUrl(urlAttr.getValue())) {
        return urlAttr.getValue();
      }
    }
  }

  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (child.getName().toLowerCase() === 'enclosure') {
      const typeAttr = child.getAttribute('type');
      const urlAttr = child.getAttribute('url');
      
      const isImageType = typeAttr && typeAttr.getValue().toLowerCase().startsWith('image/');
      if (urlAttr && (isImageType || isValidImageUrl(urlAttr.getValue()))) {
        return urlAttr.getValue();
      }
    }
  }

  return null;
}

/**
 * 記事ページにアクセスして画像（OGP / Schema.org / アイキャッチHTML）を取得する関数
 */
function fetchOgImageFromUrl(url) {
  try {
    const encodedUrl = safeUrlEncode(url);
    const options = {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8'
      }
    };
    const res = UrlFetchApp.fetch(encodedUrl, options);
    if (res.getResponseCode() !== 200) return null;

    const html = res.getContentText();

    // 優先度1: standard og:image / twitter:image
    const ogMatch = html.match(/<meta[^>]+(?:property|name)=["'](?:og:image|twitter:image)["'][^>]+content=["']([^"']+)["']/i) ||
                    html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+(?:property|name)=["'](?:og:image|twitter:image)["']/i);

    if (ogMatch && ogMatch[1]) {
      let ogUrl = ogMatch[1].trim();
      if (isValidImageUrl(ogUrl)) return ogUrl;
    }

    // 優先度2: Schema.org (itemprop="image" 内の meta itemprop="url") ※メガニケ速報対応
    const schemaMatch = html.match(/itemprop=["']image["'][\s\S]*?<meta[^>]+itemprop=["']url["'][^>]+content=["']([^"']+)["']/i) ||
                        html.match(/<meta[^>]+itemprop=["']url["'][^>]+content=["']([^"']+)["']/i);

    if (schemaMatch && schemaMatch[1]) {
      let schemaUrl = schemaMatch[1].trim();
      if (isValidImageUrl(schemaUrl)) return schemaUrl;
    }

    // 優先度3: アイキャッチ画像クラス（eye-catch-image / wp-post-image 等）が付与された img タグ
    const eyeCatchMatch = html.match(/<img[^>]+class=["'][^"']*(?:eye-catch|wp-post-image|featured-image)[^"']*["'][^>]+(?:src|srcset)=["']([^"'\s>]+)["']/i);

    if (eyeCatchMatch && eyeCatchMatch[1]) {
      let imgUrl = eyeCatchMatch[1].split(',')[0].trim().split(' ')[0];
      if (imgUrl.startsWith('//')) imgUrl = 'https:' + imgUrl;
      if (isValidImageUrl(imgUrl)) return imgUrl;
    }

    // 優先度4: link rel="image_src"
    const linkSrcMatch = html.match(/<link[^>]+rel=["']image_src["'][^>]+href=["']([^"']+)["']/i);
    if (linkSrcMatch && linkSrcMatch[1]) {
      let linkUrl = linkSrcMatch[1].trim();
      if (isValidImageUrl(linkUrl)) return linkUrl;
    }

  } catch (e) {
    console.warn(`[Web画像取得失敗] ${url}: ${e.message}`);
  }
  return null;
}

/**
 * 有効な画像URLか判定（動画・コメントキャラ画像・アイコン・ロゴ・広告・SVGを厳格除外）
 */
function isValidImageUrl(url) {
  if (!url || typeof url !== 'string' || (!url.startsWith('http://') && !url.startsWith('https://'))) {
    return false;
  }

  const lower = url.toLowerCase();
  const cleanPath = lower.split('?')[0].split('#')[0];

  // 1. 動画ファイルおよび非対応形式（SVG / GIF / DataURI）を除外
  const invalidExtensions = ['.mp4', '.webm', '.mkv', '.avi', '.mov', '.flv', '.wmv', '.m4v', '.svg'];
  for (let i = 0; i < invalidExtensions.length; i++) {
    if (cleanPath.endsWith(invalidExtensions[i])) return false;
  }

  // 2. コメントアイコン・装飾画像・広告・固定看板用画像を厳格除外
  const ignoreKeywords = [
    'comment', 'chara', 'icon', 'avatar', 'res_', 'thumb_comment', // コメント領域・キャラ画像
    'amazon.com', 'amazon-adsystem.com', 'm.media-amazon.com', 'ssl-images-amazon.com',
    'pochipp', 'pochipp-logo', 'plugins/pochipp',
    '32381cb2.jpg',       // サイト看板ロゴ
    '/smilies/', 'emoji', 'counter', 'facebook.com', 'twitter.com', 
    'line.me', 'hatena', 'share', 'clear.gif', 
    'blank.gif', 'pixel', 'ad_banner', 'logo_publisher'
  ];

  for (let i = 0; i < ignoreKeywords.length; i++) {
    if (lower.includes(ignoreKeywords[i])) return false;
  }

  // 3. 静止画フォーマットの判定
  return (
    lower.includes('wp-content/uploads') ||
    cleanPath.endsWith('.jpg') ||
    cleanPath.endsWith('.jpeg') ||
    cleanPath.endsWith('.png') ||
    cleanPath.endsWith('.webp')
  );
}

/**
 * Discord通知機能（画像を添付ファイルとして直接送信することで直リンク拒否を回避）
 */
function sendDiscordEmbedNotification(webhookUrl, title, link, pubTime, siteName, customIconUrl, defaultIconUrl, hexColorStr, imageUrl) {
  const safeLink = safeUrlEncode(link);

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
      icon_url: footerIcon ? safeUrlEncode(footerIcon) : undefined
    },
    timestamp: isoTimestamp
  };

  const payload = {
    content: title,
    embeds: [embed]
  };

  if (siteName && siteName.trim() !== '') {
    payload.username = siteName.substring(0, 80);
  }

  if (customIconUrl && (customIconUrl.startsWith('http://') || customIconUrl.startsWith('https://'))) {
    payload.avatar_url = safeUrlEncode(customIconUrl);
  }

  // 画像の処理：GAS側で画像をダウンロードしてDiscordに添付送信する
  let imageBlob = null;
  if (imageUrl) {
    try {
      const imgRes = UrlFetchApp.fetch(imageUrl, {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
      });
      if (imgRes.getResponseCode() === 200) {
        imageBlob = imgRes.getBlob().setName('embed_image.png');
        // attachment:// スキーマを使って添付画像をEmbedに指定
        embed.image = { url: 'attachment://embed_image.png' };
        console.log(`  └ [画像ダウンロード成功]: 添付ファイルとして送信します`);
      } else {
        console.warn(`  └ [画像ダウンロード失敗] HTTP ${imgRes.getResponseCode()}: 直リンクフォールバックを試みます`);
        embed.image = { url: safeUrlEncode(imageUrl) };
      }
    } catch (e) {
      console.warn(`  └ [画像取得例外]: ${e.message}`);
      embed.image = { url: safeUrlEncode(imageUrl) };
    }
  }

  // リクエストの作成（画像添付がある場合は multipart/form-data）
  let options = {};
  if (imageBlob) {
    const formData = {
      'payload_json': Utilities.newBlob(JSON.stringify(payload), 'application/json'),
      'files[0]': imageBlob
    };
    options = {
      method: 'post',
      payload: formData,
      muteHttpExceptions: true
    };
  } else {
    options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
  }

  const response = UrlFetchApp.fetch(webhookUrl, options);
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    console.error(`[Discord送信失敗] HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
  }
}

/**
 * 日本語・特殊文字・二重エンコードに対応した高精度URLエンコード関数
 */
function safeUrlEncode(rawUrl) {
  if (!rawUrl || typeof rawUrl !== 'string') return '';

  let clean = rawUrl.replace(/[\r\n\t]/g, '').trim();
  if (clean.startsWith('//')) clean = 'https:' + clean;

  // HTMLエンティティ(&amp;等)を完全解除
  clean = decodeHtmlEntitiesFully(clean);

  // すでにパーセントエンコード済みの場合は1回完全解凍
  try {
    let decoded = clean;
    while (decoded.includes('%')) {
      const prev = decoded;
      decoded = decodeURIComponent(decoded);
      if (prev === decoded) break;
    }
    clean = decoded;
  } catch (e) {
    // 解凍失敗時はそのまま進む
  }

  // マルチバイト文字（日本語など）および特殊文字を安全に置換・エンコード
  return clean.replace(/[^\x00-\x7F]/g, function(c) {
    return encodeURIComponent(c);
  }).replace(/ /g, '%20');
}

/* ヘルパー関数群 */
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