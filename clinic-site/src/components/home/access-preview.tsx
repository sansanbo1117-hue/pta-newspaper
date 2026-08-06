import Link from "next/link";
import { ArrowRight } from "lucide-react";

import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";
import { AccessMap } from "@/components/shared/access-map";

export function AccessPreview() {
  return (
    <section className="py-14 sm:py-20">
      <Container className="flex flex-col gap-10">
        <SectionHeading
          eyebrow="ACCESS"
          title="アクセス"
          description="初診の方が迷わないよう、地図・駐車場・入口の情報を大きく掲載しています。"
        />

        <AccessMap />

        <div className="flex justify-center">
          <Link
            href="/access"
            className="inline-flex items-center gap-1.5 text-base font-bold text-brand-700 hover:underline"
          >
            アクセス詳細を見る
            <ArrowRight className="size-4" aria-hidden />
          </Link>
        </div>
      </Container>
    </section>
  );
}
