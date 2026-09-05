/**
 * 処理の設定および除外キーワードの定義ファイル
 */
const CONFIG = {
  // 1回の実行で送信する最大記事数（Discordレートリミット対策）
  MAX_POSTS_PER_RUN: 5,

  // デフォルトの埋め込みカラー（16進数数値）
  DEFAULT_COLOR: 0x3498db,

  // 除外する動画・非対応拡張子
  INVALID_EXTENSIONS: [
    '.mp4', '.webm', '.mkv', '.avi', '.mov', '.flv', '.wmv', '.m4v', '.svg', '.gif'
  ],

  // 有効とみなす画像拡張子
  VALID_IMAGE_EXTENSIONS: [
    '.jpg', '.jpeg', '.png', '.webp'
  ],

  // コメントアイコン・装飾画像・広告・ロゴ等の除外キーワード
  IGNORE_KEYWORDS: [
    'comment', 'chara', 'icon', 'avatar', 'res_', 'thumb_comment',
    'amazon.com', 'amazon-adsystem.com', 'm.media-amazon.com', 'ssl-images-amazon.com',
    'pochipp', 'pochipp-logo', 'plugins/pochipp',
    '/smilies/', 'emoji', 'counter', 'facebook.com', 'twitter.com',
    'line.me', 'hatena', 'share', 'clear.gif', 'default.jpg',
    'blank.gif', 'pixel', 'ad_banner', 'logo_publisher'
  ]
};
