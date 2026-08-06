import {
  CalendarClock,
  Accessibility,
  Car,
  MessagesSquare,
  ScanLine,
  Users,
} from "lucide-react";

import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";

const features = [
  {
    icon: CalendarClock,
    title: "待ち時間を減らすWeb予約",
    body: "24時間いつでもスマホから予約可能。順番の目安が分かるので、院内で長く待つ心配がありません。",
  },
  {
    icon: Accessibility,
    title: "高齢の方にもやさしい院内",
    body: "段差の少ないバリアフリー設計。車椅子・杖をお使いの方もスタッフがサポートします。",
  },
  {
    icon: Car,
    title: "駐車場完備",
    body: "お車での通院がしやすいよう駐車場をご用意。入口までの動線も分かりやすくご案内します。",
  },
  {
    icon: ScanLine,
    title: "院内でのレントゲン検査",
    body: "院内にレントゲン設備があり、その日のうちに検査から診断まで行えます。",
  },
  {
    icon: MessagesSquare,
    title: "分かりやすい説明",
    body: "画像や模型を使いながら、専門用語をかみ砕いて丁寧にご説明します。",
  },
  {
    icon: Users,
    title: "経験豊富な医師・スタッフ",
    body: "整形外科・スポーツ医学の専門知識を持つ医師と、理学療法士が連携して治療にあたります。",
  },
] as const;

export function Features() {
  return (
    <section className="bg-muted py-14 sm:py-20">
      <Container className="flex flex-col gap-10">
        <SectionHeading
          eyebrow="FEATURES"
          title="当院の特徴"
          description="患者さまにも、私たちスタッフにも無理のない、安心して通い続けられるクリニックを目指しています。"
        />

        <ul className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {features.map(({ icon: Icon, title, body }) => (
            <li
              key={title}
              className="flex flex-col gap-3 rounded-card bg-white p-6 shadow-soft"
            >
              <span className="flex size-12 items-center justify-center rounded-full bg-brand-50 text-brand-700">
                <Icon className="size-6" aria-hidden />
              </span>
              <h3 className="text-lg font-bold leading-snug">{title}</h3>
              <p className="text-sm leading-relaxed text-muted-foreground">
                {body}
              </p>
            </li>
          ))}
        </ul>
      </Container>
    </section>
  );
}
