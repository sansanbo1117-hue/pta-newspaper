# 公開前チェックリスト

このプロジェクトは、既存サイト（www.aocli.com）へのネットワークアクセスが制限された環境で
構築されたため、実データの代わりにプレースホルダーを使用している箇所があります。
公開前に以下を対応してください。

## 1. 実コンテンツへの差し替え

- `src/data/site.ts`：クリニック名の英語表記、電話番号、FAX番号、住所、最寄り駅・バス・駐車場情報、
  Googleマップ埋め込みURL（座標を実際の住所に更新）
- `src/data/doctors.ts`：院長・医師の氏名、経歴、専門分野、資格、メッセージ、写真
- `src/data/services.ts`：診療科目の内容、初診の流れの文言
- `src/data/faq.ts`：実際に電話で多い質問への差し替え・追加
- `src/data/news.ts`：休診案内・お知らせの実データ（将来的にはCMS/管理画面からの入稿に置き換え）
- `src/data/jobs.ts`：募集職種・条件
- `src/app/about/page.tsx`：ごあいさつ文、医院概要
- `src/app/privacy/page.tsx`：プライバシーポリシーの内容（顧問弁護士等の確認を推奨）
- 写真素材：ヒーロー画像、院内紹介（待合室・診察室・レントゲン室・リハビリ室等）、
  駐車場・入口の写真、医師の顔写真（現在はすべてアイコンのプレースホルダー）

## 2. バックエンド連携（現在はプロトタイプ実装）

- `src/app/contact/actions.ts`：お問い合わせフォームの送信処理。現状はコンソール出力のみ。
  メール送信サービス（Resend, SendGrid など）やCRM/電子カルテ連携を実装すること。
- `src/app/reserve/actions.ts`：Web予約フォームの送信処理。現状は「スタッフが内容を確認し
  折り返し電話で確定する」リクエスト型として実装（コンソール出力のみ）。
  将来的にリアルタイム予約システムに接続する場合も、フォームUI側は変更不要な設計にしてある。
- お知らせ機能：現状は `src/data/news.ts` を直接編集する運用。将来的に管理画面（CMS）を
  追加する場合は、この配列と同じ形（slug/date/category/title/body）のデータソースに
  差し替えれば既存のUIがそのまま使える。

## 3. 公開直前に戻す設定

実データ反映・内容確認が完了したら、検索エンジンのクロールを許可する：

- `src/app/layout.tsx` の `metadata.robots` を `{ index: true, follow: true }` に変更
- `src/app/robots.ts` の `disallow: "/"` を `allow: "/"` に変更

## 4. その他

- `src/data/site.ts` の `mapEmbedSrc` / `mapLinkHref` を実際の住所・Google Maps URLに更新
- OGP画像（`src/app/opengraph-image` 等）の追加を検討
- Google Search Console / Google Business Profile との連携
