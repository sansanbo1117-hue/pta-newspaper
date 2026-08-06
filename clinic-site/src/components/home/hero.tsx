import Link from "next/link";
import { Phone, CalendarCheck, ShieldCheck } from "lucide-react";

import { siteConfig } from "@/data/site";
import { Button } from "@/components/ui/button";
import { Container } from "@/components/shared/container";

export function Hero() {
  return (
    <section className="relative overflow-hidden bg-gradient-to-b from-brand-50 via-white to-white">
      <div
        aria-hidden
        className="pointer-events-none absolute -right-24 -top-24 size-96 rounded-full bg-brand-100/70 blur-3xl"
      />
      <div
        aria-hidden
        className="pointer-events-none absolute -left-32 top-40 size-80 rounded-full bg-brand-200/50 blur-3xl"
      />

      <Container className="relative grid gap-10 py-12 sm:py-16 lg:grid-cols-2 lg:items-center lg:py-24">
        <div className="flex flex-col items-start gap-6">
          <span className="inline-flex items-center gap-2 rounded-full bg-white px-4 py-1.5 text-sm font-bold text-brand-700 shadow-soft ring-1 ring-inset ring-brand-100">
            <ShieldCheck className="size-4" aria-hidden />
            地域の「動く」を支える整形外科
          </span>
          <h1 className="text-3xl font-bold leading-snug text-foreground sm:text-4xl lg:text-5xl">
            痛みを我慢しない。
            <br />
            <span className="text-brand-700">{siteConfig.name}</span>
          </h1>
          <p className="max-w-xl text-base leading-relaxed text-muted-foreground sm:text-lg">
            骨折・捻挫・腰痛・肩こりからスポーツ外傷、リハビリテーションまで。
            お子さまからご高齢の方まで、安心して通える地域のかかりつけ医を目指します。
          </p>

          <div className="flex w-full flex-col gap-3 sm:w-auto sm:flex-row">
            <Button asChild size="lg">
              <a href={siteConfig.telHref}>
                <Phone /> {siteConfig.telDisplay}
              </a>
            </Button>
            <Button asChild size="lg" variant="warm">
              <Link href="/reserve">
                <CalendarCheck /> Web予約はこちら
              </Link>
            </Button>
          </div>
        </div>

        <div
          role="img"
          aria-label="明るく清潔感のあるクリニックの受付イメージ"
          className="relative flex aspect-4/3 items-center justify-center overflow-hidden rounded-card bg-gradient-to-br from-brand-500 to-brand-700 shadow-soft-lg"
        >
          <div className="absolute inset-0 bg-[radial-gradient(circle_at_top_right,rgba(255,255,255,0.25),transparent_55%)]" />
          <div className="relative flex flex-col items-center gap-3 px-8 text-center text-white">
            <ShieldCheck className="size-16" aria-hidden />
            <p className="text-lg font-bold">
              院内・受付の写真は準備中です
            </p>
            <p className="text-sm text-white/80">
              公開までに実際の外観・受付・診察室の写真に差し替えます
            </p>
          </div>
        </div>
      </Container>
    </section>
  );
}
