import type { Metadata } from "next";
import { AlertCircle, Clock } from "lucide-react";

import { clinicHours, closedNote, emergencyNote, siteConfig } from "@/data/site";
import { getSortedNews, newsCategoryColor } from "@/data/news";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";

export const metadata: Metadata = {
  title: "診療時間",
  description: `${siteConfig.name}の診療時間・休診日のご案内です。`,
};

export default function HoursPage() {
  const hourChanges = getSortedNews().filter(
    (n) => n.category === "休診案内" || n.category === "診療時間変更"
  );

  return (
    <>
      <PageHeader title="診療時間" lead="診療時間・休診日のご案内です。" />

      <section className="py-12 sm:py-16">
        <Container className="flex flex-col gap-8">
          <SectionHeading align="left" title="通常の診療時間" />

          <div className="overflow-x-auto rounded-card border border-border shadow-soft">
            <table className="w-full min-w-[480px] text-center">
              <thead>
                <tr className="bg-brand-600 text-white">
                  <th scope="col" className="px-4 py-3 text-left font-bold">
                    曜日
                  </th>
                  <th scope="col" className="px-4 py-3 font-bold">
                    午前
                  </th>
                  <th scope="col" className="px-4 py-3 font-bold">
                    午後
                  </th>
                </tr>
              </thead>
              <tbody>
                {clinicHours.map((row) => (
                  <tr key={row.label} className="border-b border-border last:border-b-0 odd:bg-muted/60">
                    <th scope="row" className="px-4 py-3 text-left font-bold">
                      {row.label}
                    </th>
                    <td className="px-4 py-3 text-muted-foreground">{row.am}</td>
                    <td className="px-4 py-3 text-muted-foreground">{row.pm}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="flex items-start gap-3 rounded-card bg-muted p-5">
            <Clock className="mt-0.5 size-5 shrink-0 text-brand-600" aria-hidden />
            <div className="text-sm leading-relaxed text-muted-foreground">
              <p>受付は診療終了の30分前までにお越しください。</p>
              <p>{closedNote}</p>
            </div>
          </div>

          <div className="flex items-start gap-3 rounded-card border border-amber-200 bg-amber-50 p-5">
            <AlertCircle className="mt-0.5 size-5 shrink-0 text-amber-600" aria-hidden />
            <p className="text-sm leading-relaxed text-amber-800">{emergencyNote}</p>
          </div>
        </Container>
      </section>

      {hourChanges.length > 0 ? (
        <section className="bg-muted py-12 sm:py-16">
          <Container className="flex flex-col gap-6">
            <SectionHeading
              align="left"
              title="臨時休診・診療時間の変更"
              description="お盆・年末年始・臨時休診などの最新情報です。"
            />
            <ul className="flex flex-col gap-3">
              {hourChanges.map((item) => (
                <li
                  key={item.slug}
                  className="flex flex-col gap-2 rounded-card bg-white p-5 shadow-soft sm:flex-row sm:items-center sm:gap-4"
                >
                  <span
                    className={`inline-flex w-fit items-center rounded-full px-3 py-1 text-xs font-bold ${newsCategoryColor[item.category]}`}
                  >
                    {item.category}
                  </span>
                  <a
                    href={`/news/${item.slug}`}
                    className="font-bold text-foreground hover:text-brand-700 hover:underline"
                  >
                    {item.title}
                  </a>
                </li>
              ))}
            </ul>
          </Container>
        </section>
      ) : null}
    </>
  );
}
