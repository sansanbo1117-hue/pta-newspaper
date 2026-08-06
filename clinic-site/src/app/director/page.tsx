import type { Metadata } from "next";
import { Stethoscope } from "lucide-react";

import { doctors } from "@/data/doctors";
import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { Badge } from "@/components/ui/badge";

export const metadata: Metadata = {
  title: "院長紹介",
  description: `${siteConfig.name}の院長・医師のプロフィールをご紹介します。`,
};

export default function DirectorPage() {
  return (
    <>
      <PageHeader
        title="院長紹介"
        breadcrumbLabel="院長紹介"
        lead="当院で診療にあたる院長・医師をご紹介します。専門分野や経歴、患者さまへのメッセージをご覧ください。"
      />

      <section className="py-12 sm:py-16">
        <Container className="flex flex-col gap-10">
          {doctors.map((doctor) => (
            <article
              key={doctor.slug}
              id={doctor.slug}
              className="grid gap-8 rounded-card border border-border bg-white p-6 shadow-soft sm:p-10 lg:grid-cols-[220px_1fr]"
            >
              <div className="flex flex-col items-center gap-3 text-center lg:items-start lg:text-left">
                <div
                  role="img"
                  aria-label={`${doctor.name}${doctor.role}の写真（準備中）`}
                  className="flex size-40 items-center justify-center rounded-full bg-brand-50 text-brand-700"
                >
                  <Stethoscope className="size-16" aria-hidden />
                </div>
                {doctor.isDirector ? <Badge>院長</Badge> : null}
              </div>

              <div className="flex flex-col gap-5">
                <div>
                  <p className="text-sm font-bold text-brand-600">{doctor.role}</p>
                  <h2 className="text-2xl font-bold">
                    {doctor.name}
                    <span className="ml-2 text-base font-normal text-muted-foreground">
                      {doctor.nameKana}
                    </span>
                  </h2>
                </div>

                <p className="leading-relaxed text-muted-foreground">
                  {doctor.message}
                </p>

                <div className="grid gap-6 sm:grid-cols-2">
                  <div>
                    <h3 className="mb-2 font-bold">専門分野</h3>
                    <ul className="flex flex-wrap gap-2">
                      {doctor.specialties.map((s) => (
                        <li key={s}>
                          <Badge variant="outline">{s}</Badge>
                        </li>
                      ))}
                    </ul>
                  </div>
                  <div>
                    <h3 className="mb-2 font-bold">資格</h3>
                    <ul className="list-inside list-disc text-sm leading-relaxed text-muted-foreground">
                      {doctor.qualifications.map((q) => (
                        <li key={q}>{q}</li>
                      ))}
                    </ul>
                  </div>
                </div>

                <div>
                  <h3 className="mb-2 font-bold">経歴</h3>
                  <ul className="list-inside list-disc text-sm leading-relaxed text-muted-foreground">
                    {doctor.career.map((c) => (
                      <li key={c}>{c}</li>
                    ))}
                  </ul>
                </div>
              </div>
            </article>
          ))}
        </Container>
      </section>
    </>
  );
}
