import type { Metadata } from "next";
import {
  Armchair,
  Dumbbell,
  DoorOpen,
  ScanLine,
  Stethoscope,
  Users,
} from "lucide-react";

import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";

export const metadata: Metadata = {
  title: "院内紹介",
  description: `${siteConfig.name}の院内・設備をご紹介します。`,
};

const rooms = [
  { icon: DoorOpen, title: "エントランス", body: "段差の少ないバリアフリー設計の入口です。" },
  { icon: Armchair, title: "待合室", body: "ゆったりとお過ごしいただける待合スペースです。" },
  { icon: Users, title: "受付", body: "スタッフが丁寧にご案内いたします。" },
  { icon: Stethoscope, title: "診察室", body: "落ち着いてご相談いただける診察室です。" },
  { icon: ScanLine, title: "レントゲン室", body: "院内で検査から診断まで行えます。" },
  { icon: Dumbbell, title: "リハビリテーション室", body: "理学療法士が個別にリハビリを行います。" },
];

export default function FacilityPage() {
  return (
    <>
      <PageHeader
        title="院内紹介"
        lead="初めての方でも安心してお越しいただけるよう、院内の様子をご紹介します。"
      />

      <section className="py-12 sm:py-16">
        <Container className="flex flex-col gap-8">
          <SectionHeading
            align="left"
            title="院内フロア"
            description="実際の写真は準備が整い次第、順次差し替えます。"
          />
          <ul className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
            {rooms.map(({ icon: Icon, title, body }) => (
              <li key={title}>
                <figure className="overflow-hidden rounded-card bg-white shadow-soft">
                  <div className="flex aspect-4/3 flex-col items-center justify-center gap-2 bg-brand-50 text-brand-700">
                    <Icon className="size-10" aria-hidden />
                    <span className="text-sm font-bold">写真準備中</span>
                  </div>
                  <figcaption className="p-5">
                    <p className="font-bold">{title}</p>
                    <p className="mt-1 text-sm leading-relaxed text-muted-foreground">
                      {body}
                    </p>
                  </figcaption>
                </figure>
              </li>
            ))}
          </ul>
        </Container>
      </section>

      <section className="bg-muted py-12 sm:py-16">
        <Container className="flex flex-col gap-4">
          <SectionHeading align="left" title="バリアフリー・設備について" />
          <ul className="list-inside list-disc space-y-2 text-sm leading-relaxed text-muted-foreground sm:text-base">
            <li>院内は段差が少なく、車椅子・ベビーカーでもご利用いただけます。</li>
            <li>多目的トイレを設置しています。</li>
            <li>お子さま連れの方も安心のキッズスペースをご用意しています。</li>
          </ul>
        </Container>
      </section>
    </>
  );
}
