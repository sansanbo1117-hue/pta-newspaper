import Link from "next/link";
import { Phone, Home } from "lucide-react";

import { siteConfig } from "@/data/site";
import { Container } from "@/components/shared/container";
import { Button } from "@/components/ui/button";

export default function NotFound() {
  return (
    <section className="py-20 sm:py-28">
      <Container className="flex flex-col items-center gap-6 text-center">
        <p className="text-sm font-bold text-brand-600">404</p>
        <h1 className="text-2xl font-bold sm:text-3xl">
          ページが見つかりませんでした
        </h1>
        <p className="max-w-md leading-relaxed text-muted-foreground">
          お探しのページは移動または削除された可能性があります。トップページから改めてお探しください。
        </p>
        <div className="flex flex-col gap-3 sm:flex-row">
          <Button asChild size="lg">
            <Link href="/">
              <Home /> トップページへ戻る
            </Link>
          </Button>
          <Button asChild size="lg" variant="secondary">
            <a href={siteConfig.telHref}>
              <Phone /> {siteConfig.telDisplay}
            </a>
          </Button>
        </div>
      </Container>
    </section>
  );
}
