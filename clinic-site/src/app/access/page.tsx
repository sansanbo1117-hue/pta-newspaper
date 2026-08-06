import type { Metadata } from "next";

import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { AccessMap } from "@/components/shared/access-map";

export const metadata: Metadata = {
  title: "アクセス",
  description: `${siteConfig.name}への地図・駐車場・交通手段のご案内です。`,
};

export default function AccessPage() {
  return (
    <>
      <PageHeader
        title="アクセス"
        lead="初診の方にも迷わずお越しいただけるよう、地図・駐車場・入口の情報をご案内します。"
      />

      <section className="py-12 sm:py-16">
        <Container>
          <AccessMap />
        </Container>
      </section>
    </>
  );
}
