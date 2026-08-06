import type { Metadata } from "next";
import { Phone, ShieldCheck } from "lucide-react";

import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { ReservationForm } from "@/components/reserve/reservation-form";

export const metadata: Metadata = {
  title: "Web予約",
  description: `${siteConfig.name}のWeb予約ページです。初診・再診の予約、変更・キャンセルをリクエストできます。`,
};

export default function ReservePage() {
  return (
    <>
      <PageHeader
        title="Web予約"
        lead="24時間いつでも、初診・再診のご予約、予約の変更・キャンセルをリクエストいただけます。"
      />

      <section className="py-12 sm:py-16">
        <Container className="grid gap-10 lg:grid-cols-[1fr_1.4fr]">
          <div className="flex flex-col gap-4">
            <div className="flex items-start gap-3 rounded-card bg-brand-50 p-6 ring-1 ring-inset ring-brand-100">
              <ShieldCheck className="mt-0.5 size-6 shrink-0 text-brand-700" aria-hidden />
              <div className="text-sm leading-relaxed text-brand-900">
                <p className="font-bold">ご予約の流れ</p>
                <p className="mt-1">
                  フォーム送信後、内容をスタッフが確認し、お電話にてご予約を確定いたします。
                  お急ぎの場合はお電話でも承ります。
                </p>
              </div>
            </div>
            <div className="rounded-card border border-border bg-white p-6 shadow-soft">
              <p className="font-bold">お急ぎの場合はお電話で</p>
              <a
                href={siteConfig.telHref}
                className="mt-2 flex items-center gap-2 text-xl font-bold text-brand-700"
              >
                <Phone className="size-5" aria-hidden />
                {siteConfig.telDisplay}
              </a>
            </div>
            <p className="text-sm leading-relaxed text-muted-foreground">
              症状によっては当日のご案内が難しい場合があります。あらかじめご了承ください。
            </p>
          </div>

          <div className="rounded-card border border-border bg-white p-6 shadow-soft sm:p-8">
            <ReservationForm />
          </div>
        </Container>
      </section>
    </>
  );
}
