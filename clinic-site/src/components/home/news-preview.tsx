import Link from "next/link";
import { ArrowRight } from "lucide-react";

import { getSortedNews, newsCategoryColor } from "@/data/news";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";

function formatDate(iso: string) {
  const d = new Date(iso);
  return `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, "0")}.${String(
    d.getDate()
  ).padStart(2, "0")}`;
}

export function NewsPreview() {
  const news = getSortedNews().slice(0, 3);

  return (
    <section className="bg-muted py-14 sm:py-20">
      <Container className="flex flex-col gap-10">
        <SectionHeading eyebrow="NEWS" title="お知らせ" />

        <ul className="flex flex-col divide-y divide-border rounded-card bg-white shadow-soft">
          {news.map((item) => (
            <li key={item.slug}>
              <Link
                href={`/news/${item.slug}`}
                className="flex flex-col gap-2 p-5 hover:bg-muted sm:flex-row sm:items-center sm:gap-6 sm:p-6"
              >
                <span className="text-sm font-bold text-muted-foreground sm:w-28">
                  {formatDate(item.date)}
                </span>
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

        <div className="flex justify-center">
          <Link
            href="/news"
            className="inline-flex items-center gap-1.5 text-base font-bold text-brand-700 hover:underline"
          >
            お知らせ一覧を見る
            <ArrowRight className="size-4" aria-hidden />
          </Link>
        </div>
      </Container>
    </section>
  );
}
