import type { Metadata } from "next";
import { Mail, Phone } from "lucide-react";

import { jobPostings } from "@/data/jobs";
import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";

export const metadata: Metadata = {
  title: "採用情報",
  description: `${siteConfig.name}の採用情報・募集職種のご案内です。`,
};

export default function RecruitPage() {
  return (
    <>
      <PageHeader
        title="採用情報"
        lead="一緒に地域の患者さまを支えてくださるスタッフを募集しています。"
      />

      <section className="py-12 sm:py-16">
        <Container className="flex flex-col gap-6">
          <SectionHeading align="left" title="当院で働く魅力" />
          <ul className="grid gap-4 sm:grid-cols-3">
            <li className="rounded-card bg-white p-5 shadow-soft">
              <p className="font-bold">研修・教育制度</p>
              <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                経験の浅い方も安心して働けるよう、院内研修や学会参加をサポートします。
              </p>
            </li>
            <li className="rounded-card bg-white p-5 shadow-soft">
              <p className="font-bold">働きやすい環境</p>
              <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                シフトの相談がしやすく、ライフステージに合わせて働き続けられます。
              </p>
            </li>
            <li className="rounded-card bg-white p-5 shadow-soft">
              <p className="font-bold">チーム医療</p>
              <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                医師・看護師・理学療法士・事務スタッフが連携し、風通しの良い職場です。
              </p>
            </li>
          </ul>
        </Container>
      </section>

      <section className="bg-muted py-12 sm:py-16">
        <Container className="flex flex-col gap-6">
          <SectionHeading align="left" title="募集職種" />
          <ul className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
            {jobPostings.map((job) => (
              <li
                key={job.slug}
                className="flex flex-col gap-3 rounded-card bg-white p-6 shadow-soft"
              >
                <div className="flex items-center justify-between gap-2">
                  <h3 className="text-lg font-bold">{job.title}</h3>
                  <Badge variant="outline">{job.employmentType}</Badge>
                </div>
                <p className="text-sm leading-relaxed text-muted-foreground">
                  {job.summary}
                </p>
                <dl className="mt-2 space-y-1.5 text-sm">
                  <div className="flex gap-2">
                    <dt className="w-20 shrink-0 font-bold text-foreground">応募資格</dt>
                    <dd className="text-muted-foreground">
                      {job.requirements.join("、")}
                    </dd>
                  </div>
                  <div className="flex gap-2">
                    <dt className="w-20 shrink-0 font-bold text-foreground">勤務時間</dt>
                    <dd className="text-muted-foreground">{job.workingHours}</dd>
                  </div>
                  <div className="flex gap-2">
                    <dt className="w-20 shrink-0 font-bold text-foreground">給与</dt>
                    <dd className="text-muted-foreground">{job.salary}</dd>
                  </div>
                  <div className="flex gap-2">
                    <dt className="w-20 shrink-0 font-bold text-foreground">待遇</dt>
                    <dd className="text-muted-foreground">
                      {job.benefits.join("、")}
                    </dd>
                  </div>
                </dl>
              </li>
            ))}
          </ul>
        </Container>
      </section>

      <section className="py-12 sm:py-16">
        <Container className="flex flex-col items-start gap-4 rounded-card bg-brand-600 p-8 text-white sm:flex-row sm:items-center sm:justify-between">
          <div>
            <h2 className="text-xl font-bold">ご応募・お問い合わせ</h2>
            <p className="mt-1 text-sm text-white/90">
              まずはお電話またはお問い合わせフォームよりご連絡ください。
            </p>
          </div>
          <div className="flex flex-col gap-3 sm:flex-row">
            <Button asChild variant="secondary">
              <a href={siteConfig.telHref}>
                <Phone /> {siteConfig.telDisplay}
              </a>
            </Button>
            <Button asChild variant="warm">
              <a href="/contact">
                <Mail /> お問い合わせフォーム
              </a>
            </Button>
          </div>
        </Container>
      </section>
    </>
  );
}
