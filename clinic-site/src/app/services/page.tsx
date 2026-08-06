import type { Metadata } from "next";
import Link from "next/link";
import { CalendarCheck } from "lucide-react";

import { services } from "@/data/services";
import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";
import { FirstVisitSteps } from "@/components/shared/first-visit-steps";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";

export const metadata: Metadata = {
  title: "診療案内",
  description: `${siteConfig.name}の診療科目・初診の流れについてご案内します。`,
};

export default function ServicesPage() {
  return (
    <>
      <PageHeader
        title="診療案内"
        lead="骨折・スポーツ外傷から慢性的な腰痛・肩こり、リハビリテーションまで幅広く対応しています。"
      />

      <section className="py-12 sm:py-16">
        <Container className="flex flex-col gap-8">
          <SectionHeading align="left" title="診療科目" />
          <ul className="grid gap-4 sm:grid-cols-2">
            {services.map((service) => {
              const Icon = service.icon;
              return (
                <li key={service.slug} id={service.slug}>
                  <Card className="h-full">
                    <CardHeader className="flex-row items-center gap-3 space-y-0">
                      <span className="flex size-11 shrink-0 items-center justify-center rounded-full bg-brand-50 text-brand-700">
                        <Icon className="size-5" aria-hidden />
                      </span>
                      <CardTitle>{service.title}</CardTitle>
                    </CardHeader>
                    <CardContent className="flex flex-col gap-3">
                      <p className="text-sm leading-relaxed text-muted-foreground">
                        {service.summary}
                      </p>
                      <ul className="list-inside list-disc text-sm leading-relaxed text-muted-foreground">
                        {service.details.map((d) => (
                          <li key={d}>{d}</li>
                        ))}
                      </ul>
                    </CardContent>
                  </Card>
                </li>
              );
            })}
          </ul>
        </Container>
      </section>

      <section className="bg-muted py-12 sm:py-16">
        <Container className="flex flex-col gap-8">
          <SectionHeading
            align="left"
            title="初診の流れ"
            description="ご来院から会計までの流れをご案内します。混雑状況により前後する場合があります。"
          />
          <FirstVisitSteps />
          <div className="flex flex-col items-start gap-3 rounded-card bg-white p-6 shadow-soft sm:flex-row sm:items-center sm:justify-between">
            <p className="text-sm leading-relaxed text-muted-foreground">
              Web予約をご利用いただくと受付がスムーズです。保険証をお忘れなくお持ちください。
            </p>
            <Button asChild>
              <Link href="/reserve">
                <CalendarCheck /> Web予約はこちら
              </Link>
            </Button>
          </div>
        </Container>
      </section>
    </>
  );
}
