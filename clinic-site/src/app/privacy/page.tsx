import type { Metadata } from "next";

import { siteConfig } from "@/data/site";
import { PageHeader } from "@/components/shared/page-header";
import { Container } from "@/components/shared/container";

export const metadata: Metadata = {
  title: "プライバシーポリシー",
  description: `${siteConfig.name}の個人情報保護方針です。`,
};

const sections = [
  {
    title: "個人情報の取得について",
    body: "当院は、診療、Web予約、お問い合わせ対応、採用選考などの目的のために必要な範囲で、氏名・連絡先・症状等の個人情報を適正な方法により取得します。",
  },
  {
    title: "個人情報の利用目的",
    body: "取得した個人情報は、診療・治療・リハビリテーションの提供、予約管理、お問い合わせへの対応、採用選考、法令に基づく対応のために利用し、目的の範囲を超えて利用することはありません。",
  },
  {
    title: "個人情報の第三者提供について",
    body: "法令に基づく場合を除き、ご本人の同意なく個人情報を第三者に提供することはありません。ただし、他の医療機関等への紹介など、診療上必要な場合はこの限りではありません。",
  },
  {
    title: "個人情報の安全管理",
    body: "当院は、個人情報の紛失・破壊・改ざん・漏えい等を防止するため、適切な安全管理措置を講じます。",
  },
  {
    title: "Cookie・アクセス解析について",
    body: "当サイトでは、サービス向上のためCookieやアクセス解析ツールを利用する場合があります。これらの情報から個人を特定することはありません。",
  },
  {
    title: "開示・訂正・利用停止等のご請求",
    body: "ご自身の個人情報について開示・訂正・利用停止等をご希望の場合は、お問い合わせフォームまたはお電話にてご連絡ください。",
  },
  {
    title: "お問い合わせ窓口",
    body: `${siteConfig.name}　TEL：${siteConfig.telDisplay}`,
  },
];

export default function PrivacyPolicyPage() {
  return (
    <>
      <PageHeader title="プライバシーポリシー" />

      <section className="py-12 sm:py-16">
        <Container className="mx-auto flex max-w-3xl flex-col gap-8">
          <p className="leading-relaxed text-muted-foreground">
            {siteConfig.name}（以下「当院」）は、患者さまおよびウェブサイトご利用者さまの個人情報の重要性を認識し、
            以下の方針に基づき適切に取り扱います。
          </p>
          {sections.map((s, i) => (
            <div key={s.title} className="flex flex-col gap-2">
              <h2 className="text-lg font-bold">
                {i + 1}. {s.title}
              </h2>
              <p className="leading-relaxed text-muted-foreground">{s.body}</p>
            </div>
          ))}
          <p className="text-sm text-muted-foreground">
            制定日：2026年○月○日
          </p>
        </Container>
      </section>
    </>
  );
}
