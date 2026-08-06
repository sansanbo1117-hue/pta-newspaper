import type { Metadata } from "next";

import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { FaqAccordion } from "@/components/shared/faq-accordion";

export const metadata: Metadata = {
  title: "よくある質問",
  description: `${siteConfig.name}へのお電話でよくいただくご質問をまとめました。`,
};

export default function FaqPage() {
  return (
    <>
      <PageHeader
        title="よくある質問"
        lead="お電話でよくいただくご質問をまとめました。解決しない場合はお気軽にお電話・お問い合わせフォームからご連絡ください。"
      />

      <section className="py-12 sm:py-16">
        <Container className="mx-auto max-w-3xl">
          <FaqAccordion />
        </Container>
      </section>
    </>
  );
}
