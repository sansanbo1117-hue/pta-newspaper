import Link from "next/link";
import { Phone, CalendarCheck, MapPin, Clock } from "lucide-react";

import { siteConfig, clinicHours, closedNote } from "@/data/site";
import { Container } from "@/components/shared/container";

const cards = [
  {
    icon: Phone,
    title: "電話する",
    body: siteConfig.telDisplay,
    href: siteConfig.telHref,
    cta: "電話をかける",
    tone: "bg-brand-600 text-white",
  },
  {
    icon: CalendarCheck,
    title: "Web予約",
    body: "24時間いつでも予約可能",
    href: "/reserve",
    cta: "予約画面へ",
    tone: "bg-accent-warm text-white",
  },
  {
    icon: MapPin,
    title: "アクセス",
    body: siteConfig.access.train,
    href: "/access",
    cta: "地図を見る",
    tone: "bg-white text-foreground ring-2 ring-inset ring-border",
  },
  {
    icon: Clock,
    title: "診療時間",
    body: `${clinicHours[0].am} / ${clinicHours[0].pm}`,
    href: "/hours",
    cta: "詳しく見る",
    tone: "bg-white text-foreground ring-2 ring-inset ring-border",
  },
] as const;

export function QuickActions() {
  return (
    <section aria-label="よく使う操作" className="py-8 sm:py-10">
      <Container>
        <ul className="grid grid-cols-2 gap-3 sm:gap-4 lg:grid-cols-4">
          {cards.map(({ icon: Icon, title, body, href, cta, tone }) => (
            <li key={title}>
              <Link
                href={href}
                className={`group flex h-full flex-col justify-between gap-4 rounded-card p-5 shadow-soft transition-transform hover:-translate-y-0.5 ${tone}`}
              >
                <div className="flex items-center gap-3">
                  <span
                    className={`flex size-11 shrink-0 items-center justify-center rounded-full ${
                      tone.includes("white")
                        ? "bg-brand-50 text-brand-700"
                        : "bg-white/20"
                    }`}
                  >
                    <Icon className="size-5" aria-hidden />
                  </span>
                  <span className="text-lg font-bold">{title}</span>
                </div>
                <div>
                  <p className="text-sm opacity-90">{body}</p>
                  <p className="mt-2 text-sm font-bold underline-offset-4 group-hover:underline">
                    {cta} →
                  </p>
                </div>
              </Link>
            </li>
          ))}
        </ul>
        <p className="mt-4 text-center text-sm text-muted-foreground">
          {closedNote}
        </p>
      </Container>
    </section>
  );
}
