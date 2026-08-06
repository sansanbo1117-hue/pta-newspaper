import type { Metadata } from "next";
import { HeartHandshake, Target, Users2 } from "lucide-react";

import { siteConfig, clinicHours } from "@/data/site";
import { services } from "@/data/services";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";

export const metadata: Metadata = {
  title: "医院紹介",
  description: `${siteConfig.name}の理念・沿革・概要をご紹介します。`,
};

const values = [
  {
    icon: HeartHandshake,
    title: "患者さま第一",
    body: "「話しやすい」「相談しやすい」を大切に、お一人おひとりと向き合います。",
  },
  {
    icon: Target,
    title: "早期回復への伴走",
    body: "検査・診断・治療・リハビリを一貫して行い、早い社会復帰を目指します。",
  },
  {
    icon: Users2,
    title: "地域とともに",
    body: "赤ちゃんからご高齢の方まで、地域のかかりつけ医として長く寄り添います。",
  },
];

export default function AboutPage() {
  return (
    <>
      <PageHeader
        title="医院紹介"
        lead={`${siteConfig.name}のご紹介です。当院の理念・概要についてご案内します。`}
      />

      <section className="py-12 sm:py-16">
        <Container className="flex flex-col gap-6">
          <SectionHeading align="left" title="ごあいさつ" />
          <div className="max-w-3xl space-y-4 leading-relaxed text-muted-foreground">
            <p>
              このたびは{siteConfig.name}のウェブサイトをご覧いただき、誠にありがとうございます。
              当院は、骨折・捻挫といった急なけがから、腰痛・肩こりなどの慢性的な症状、
              スポーツによる外傷、リハビリテーションまで、運動器に関わる幅広い診療を行っております。
            </p>
            <p>
              「病院に行くほどではないかもしれないけれど、気になる」という段階でも、
              お気軽にご相談いただけるクリニックでありたいと考えています。
              お子さまからご高齢の方まで、安心して通っていただけるよう、
              スタッフ一同、分かりやすい説明と丁寧な診療を心がけてまいります。
            </p>
          </div>
        </Container>
      </section>

      <section className="bg-muted py-12 sm:py-16">
        <Container className="flex flex-col gap-8">
          <SectionHeading align="left" title="当院が大切にしていること" />
          <ul className="grid gap-4 sm:grid-cols-3">
            {values.map(({ icon: Icon, title, body }) => (
              <li
                key={title}
                className="flex flex-col gap-3 rounded-card bg-white p-6 shadow-soft"
              >
                <span className="flex size-11 items-center justify-center rounded-full bg-brand-50 text-brand-700">
                  <Icon className="size-5" aria-hidden />
                </span>
                <h3 className="font-bold">{title}</h3>
                <p className="text-sm leading-relaxed text-muted-foreground">
                  {body}
                </p>
              </li>
            ))}
          </ul>
        </Container>
      </section>

      <section className="py-12 sm:py-16">
        <Container className="flex flex-col gap-8">
          <SectionHeading align="left" title="医院概要" />
          <div className="overflow-hidden rounded-card border border-border">
            <table className="w-full text-sm sm:text-base">
              <tbody>
                {[
                  ["医院名", siteConfig.name],
                  ["所在地", siteConfig.addressFull],
                  ["電話番号", siteConfig.telDisplay],
                  ["FAX番号", siteConfig.fax],
                  [
                    "診療科目",
                    services.map((s) => s.title).join("・"),
                  ],
                  [
                    "診療時間",
                    `${clinicHours[0].am} / ${clinicHours[0].pm}（詳細は「診療時間」ページをご確認ください）`,
                  ],
                ].map(([label, value]) => (
                  <tr key={label} className="border-b border-border last:border-b-0">
                    <th
                      scope="row"
                      className="w-32 bg-muted px-4 py-3 text-left align-top font-bold sm:w-48 sm:px-6"
                    >
                      {label}
                    </th>
                    <td className="px-4 py-3 leading-relaxed text-muted-foreground sm:px-6">
                      {value}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Container>
      </section>
    </>
  );
}
