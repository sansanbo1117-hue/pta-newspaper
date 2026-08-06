export type NewsCategory = "休診案内" | "診療時間変更" | "お知らせ" | "採用";

export type NewsItem = {
  slug: string;
  date: string; // ISO 8601, e.g. "2026-08-01"
  category: NewsCategory;
  title: string;
  body: string[]; // paragraphs
};

/**
 * お知らせ一覧。
 *
 * 将来的には管理画面（お知らせ管理機能）からこのデータと同じ形（slug/date/category/title/body）で
 * 追加・編集できるようにする想定。現時点ではこの配列を直接編集することで更新できる。
 * 休診・お盆・年末年始・診療時間変更などはcategoryで分類し、一覧・トップページで色分け表示する。
 */
export const newsItems: NewsItem[] = [
  {
    slug: "2026-summer-holiday",
    date: "2026-08-01",
    category: "休診案内",
    title: "夏季休診のお知らせ（8/13〜8/16）",
    body: [
      "誠に勝手ながら、8月13日（木）〜8月16日（日）は夏季休診とさせていただきます。",
      "休診期間中の急なお怪我・お痛みについては、お近くの救急医療機関をご利用ください。",
      "8月17日（月）より通常診療を再開いたします。ご不便をおかけしますが、よろしくお願いいたします。",
    ],
  },
  {
    slug: "2026-hours-change",
    date: "2026-07-10",
    category: "診療時間変更",
    title: "水曜午後の診療時間変更について",
    body: [
      "2026年8月より、水曜日の午後診療を休診とさせていただきます。",
      "ご不便をおかけいたしますが、何卒よろしくお願い申し上げます。",
    ],
  },
  {
    slug: "2026-web-reservation",
    date: "2026-06-01",
    category: "お知らせ",
    title: "Web予約サービスを開始しました",
    body: [
      "このたび、24時間いつでもご都合の良いときにご予約いただけるWeb予約サービスを開始しました。",
      "初診・再診のご予約に加え、予約の変更・キャンセルもWeb上で行えます。ぜひご利用ください。",
    ],
  },
];

export function getNewsBySlug(slug: string) {
  return newsItems.find((item) => item.slug === slug);
}

export function getSortedNews() {
  return [...newsItems].sort(
    (a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()
  );
}

export const newsCategoryColor: Record<NewsCategory, string> = {
  休診案内: "bg-rose-50 text-rose-700 ring-1 ring-inset ring-rose-200",
  診療時間変更: "bg-amber-50 text-amber-700 ring-1 ring-inset ring-amber-200",
  お知らせ: "bg-brand-50 text-brand-700 ring-1 ring-inset ring-brand-200",
  採用: "bg-emerald-50 text-emerald-700 ring-1 ring-inset ring-emerald-200",
};
