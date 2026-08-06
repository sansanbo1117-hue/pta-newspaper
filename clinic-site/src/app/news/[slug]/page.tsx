import type { Metadata } from "next";
import Link from "next/link";
import { notFound } from "next/navigation";
import { ArrowLeft } from "lucide-react";

import { getNewsBySlug, newsCategoryColor, newsItems } from "@/data/news";
import { Container } from "@/components/shared/container";
import { PageHeader } from "@/components/shared/page-header";

export function generateStaticParams() {
  return newsItems.map((item) => ({ slug: item.slug }));
}

export async function generateMetadata({
  params,
}: PageProps<"/news/[slug]">): Promise<Metadata> {
  const { slug } = await params;
  const item = getNewsBySlug(slug);
  if (!item) return {};
  return { title: item.title, description: item.body[0] };
}

function formatDate(iso: string) {
  const d = new Date(iso);
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
}

export default async function NewsDetailPage({
  params,
}: PageProps<"/news/[slug]">) {
  const { slug } = await params;
  const item = getNewsBySlug(slug);

  if (!item) {
    notFound();
  }

  return (
    <>
      <PageHeader title="お知らせ" breadcrumbLabel={item.title} />

      <section className="py-12 sm:py-16">
        <Container className="mx-auto max-w-3xl">
          <article className="rounded-card border border-border bg-white p-6 shadow-soft sm:p-10">
            <header className="mb-6 flex flex-col gap-3 border-b border-border pb-6">
              <span
                className={`inline-flex w-fit items-center rounded-full px-3 py-1 text-xs font-bold ${newsCategoryColor[item.category]}`}
              >
                {item.category}
              </span>
              <h1 className="text-2xl font-bold leading-snug sm:text-3xl">
                {item.title}
              </h1>
              <time dateTime={item.date} className="text-sm text-muted-foreground">
                {formatDate(item.date)}
              </time>
            </header>
            <div className="flex flex-col gap-4 leading-relaxed text-muted-foreground">
              {item.body.map((paragraph, i) => (
                <p key={i}>{paragraph}</p>
              ))}
            </div>
          </article>

          <Link
            href="/news"
            className="mt-6 inline-flex items-center gap-1.5 text-sm font-bold text-brand-700 hover:underline"
          >
            <ArrowLeft className="size-4" aria-hidden />
            お知らせ一覧に戻る
          </Link>
        </Container>
      </section>
    </>
  );
}
