import Link from "next/link";
import { ArrowRight, HelpCircle } from "lucide-react";

import { faqItems } from "@/data/faq";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";

export function FaqPreview() {
  const picks = faqItems.slice(0, 4);

  return (
    <section className="py-14 sm:py-20">
      <Container className="flex flex-col gap-10">
        <SectionHeading
          eyebrow="FAQ"
          title="よくあるご質問"
          description="お電話でよくいただくご質問です。お電話の前にぜひご確認ください。"
        />

        <ul className="grid gap-4 sm:grid-cols-2">
          {picks.map((item) => (
            <li
              key={item.question}
              className="flex items-start gap-3 rounded-card bg-white p-5 shadow-soft"
            >
              <HelpCircle className="mt-0.5 size-5 shrink-0 text-brand-600" aria-hidden />
              <div>
                <p className="font-bold">{item.question}</p>
                <p className="mt-1 text-sm leading-relaxed text-muted-foreground">
                  {item.answer}
                </p>
              </div>
            </li>
          ))}
        </ul>

        <div className="flex justify-center">
          <Link
            href="/faq"
            className="inline-flex items-center gap-1.5 text-base font-bold text-brand-700 hover:underline"
          >
            よくある質問をもっと見る
            <ArrowRight className="size-4" aria-hidden />
          </Link>
        </div>
      </Container>
    </section>
  );
}
