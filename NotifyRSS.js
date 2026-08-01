/**
 * Discord WebhookへEmbed（カード）形式で通知を送信する関数
 */
function sendDiscordEmbedNotification(webhookUrl, title, link, pubTime, siteName, iconUrl, hexColorStr, imageUrl) {
  const safeLink = encodeURI(link);

  // 16進数カラーコードを整数（DEC）に変換（正規表現を使わない安全な判定）
  let colorNum = 0x3498db; // デフォルト：青
  if (hexColorStr) {
    const cleanHex = String(hexColorStr).replace('#', '').trim();
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
    content: title,                 // メッセージ本文（カード上部）
    username: siteName || undefined, // Webhookの表示名にサイト名を設定
    avatar_url: iconUrl || undefined, // Webhookのアバターにサイトアイコンを設定
    embeds: [embed]
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
 * RSSアイテムから画像を抽出する関数（HTMLエンコード対応強化版）
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

  // 本文（description または content:encoded）を取得
  let description = getElementTextByNames(item, ['encoded', 'description', 'content']);
  if (description) {
    // &lt;img src="..." &gt; のようなHTMLエスケープ文字を元の <img src="..."> にデコード
    description = description
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#039;/g, "'")
      .replace(/&amp;/g, '&');

    // <img> タグの src 属性から画像URLを抽出
    const imgMatch = description.match(/<img[^>]+src=["']([^"']+)["']/i);
    if (imgMatch && imgMatch[1]) {
      // 絵文字画像（twemojiやwp-includesのスマイリーアイコン等）を除外するフィルター
      const imgSrc = imgMatch[1];
      if (!imgSrc.includes('/smilies/') && !imgSrc.includes('emoji')) {
        return imgSrc;
      }
    }
  }

  return null;
}