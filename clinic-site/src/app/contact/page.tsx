import type { Metadata } from "next";
import { Phone, Printer } from "lucide-react";

import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { ContactForm } from "@/components/contact/contact-form";

export const metadata: Metadata = {
  title: "お問い合わせ",
  description: `${siteConfig.name}へのお問い合わせフォームです。`,
};

export default function ContactPage() {
  return (
    <>
      <PageHeader
        title="お問い合わせ"
        lead="ご予約は「Web予約」ページをご利用ください。診療内容に関するご質問など、お急ぎの場合はお電話がスムーズです。"
      />

      <section className="py-12 sm:py-16">
        <Container className="grid gap-10 lg:grid-cols-[1fr_1.4fr]">
          <div className="flex flex-col gap-4">
            <div className="rounded-card bg-brand-600 p-6 text-white shadow-soft">
              <p className="text-sm font-bold text-white/80">お急ぎの場合はお電話で</p>
              <a
                href={siteConfig.telHref}
                className="mt-2 flex items-center gap-2 text-2xl font-bold"
              >
                <Phone className="size-6" aria-hidden />
                {siteConfig.telDisplay}
              </a>
              <p className="mt-3 text-sm text-white/80">受付時間は「診療時間」ページをご確認ください。</p>
            </div>
            <div className="rounded-card border border-border bg-white p-6 shadow-soft">
              <p className="flex items-center gap-2 font-bold">
                <Printer className="size-5 text-brand-600" aria-hidden />
                FAXでのお問い合わせ
              </p>
              <p className="mt-2 text-sm text-muted-foreground">{siteConfig.fax}</p>
            </div>
            <p className="text-sm leading-relaxed text-muted-foreground">
              個別の症状に関するご質問には、フォームではお答えできない場合があります。
              その際は受診のうえ、医師にご相談ください。
            </p>
          </div>

          <div className="rounded-card border border-border bg-white p-6 shadow-soft sm:p-8">
            <ContactForm />
          </div>
        </Container>
      </section>
    </>
  );
}
