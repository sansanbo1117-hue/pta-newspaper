import type { Metadata } from "next";
import Link from "next/link";

import { getSortedNews, newsCategoryColor } from "@/data/news";
import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";

export const metadata: Metadata = {
  title: "お知らせ",
  description: `${siteConfig.name}からのお知らせ・休診案内一覧です。`,
};

function formatDate(iso: string) {
  const d = new Date(iso);
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

export default function NewsPage() {
  const news = getSortedNews();

  return (
    <>
      <PageHeader title="お知らせ" lead="休診案内・診療時間の変更・当院からのお知らせを掲載しています。" />

      <section className="py-12 sm:py-16">
        <Container>
          <ul className="flex flex-col divide-y divide-border rounded-card border border-border bg-white shadow-soft">
            {news.map((item) => (
              <li key={item.slug}>
                <Link
                  href={`/news/${item.slug}`}
                  className="flex flex-col gap-2 p-5 hover:bg-muted sm:flex-row sm:items-center sm:gap-6 sm:p-6"
                >
                  <time
                    dateTime={item.date}
                    className="text-sm font-bold text-muted-foreground sm:w-36"
                  >
                    {formatDate(item.date)}
                  </time>
                  <span
                    className={`inline-flex w-fit items-center rounded-full px-3 py-1 text-xs font-bold sm:w-32 sm:justify-center ${newsCategoryColor[item.category]}`}
                  >
                    {item.category}
                  </span>
                  <span className="font-bold leading-snug text-foreground">
                    {item.title}
                  </span>
                </Link>
              </li>
            ))}
          </ul>
        </Container>
      </section>
    </>
  );
}
