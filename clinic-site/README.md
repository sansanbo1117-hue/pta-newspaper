# 青山整形外科クリニック 患者ポータル型ホームページ

Next.js (App Router) + TypeScript + Tailwind CSS + shadcn/ui スタイルのコンポーネントで構築した、
既存サイト（www.aocli.com）のリニューアル用プロジェクトです。

「見た目を綺麗にすること」ではなく、患者・病院双方に価値を提供する病院DXの第一歩となる
ホームページ兼患者ポータルを目指しています。

## セットアップ

```bash
npm install
npm run dev
```

[http://localhost:3000](http://localhost:3000) で確認できます。

## 主なディレクトリ構成

- `src/app/`：ページ（App Router）。1ディレクトリ1ルート。
- `src/components/layout/`：ヘッダー・フッター・モバイルナビなど共通レイアウト
- `src/components/home/`：トップページ専用セクション
- `src/components/shared/`：複数ページで使う共通パーツ（アクセス地図、初診の流れ、FAQ等）
- `src/components/ui/`：shadcn/ui方式の基礎UIコンポーネント（Button, Card, Accordion等）
- `src/data/`：クリニック情報・診療案内・お知らせ・FAQ・採用情報などのコンテンツデータ
- `src/app/contact/actions.ts`, `src/app/reserve/actions.ts`：フォーム送信のServer Actions

## 公開前に必ず確認すること

実データ・写真の反映、フォームのバックエンド連携、noindex設定の解除など、
公開前に対応が必要な項目は [`CONTENT-TODO.md`](./CONTENT-TODO.md) にまとめています。

## Lint / Build

```bash
npx eslint .
npm run build
```
