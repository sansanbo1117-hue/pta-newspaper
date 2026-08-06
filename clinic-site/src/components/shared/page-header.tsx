import Link from "next/link";
import { ChevronRight, Home } from "lucide-react";

import { Container } from "@/components/shared/container";

export function PageHeader({
  title,
  lead,
  breadcrumbLabel,
}: {
  title: string;
  lead?: string;
  breadcrumbLabel?: string;
}) {
  return (
    <div className="border-b border-border bg-muted">
      <Container className="flex flex-col gap-4 py-10 sm:py-14">
        <nav aria-label="パンくずリスト">
          <ol className="flex flex-wrap items-center gap-1.5 text-sm text-muted-foreground">
            <li className="flex items-center gap-1.5">
              <Link
                href="/"
                className="flex items-center gap-1 rounded hover:text-brand-700"
              >
                <Home className="size-4" aria-hidden />
                トップ
              </Link>
            </li>
            <li aria-hidden>
              <ChevronRight className="size-4" />
            </li>
            <li className="font-bold text-foreground" aria-current="page">
              {breadcrumbLabel ?? title}
            </li>
          </ol>
        </nav>
        <div className="flex flex-col gap-3">
          <h1 className="text-3xl font-bold leading-snug text-foreground sm:text-4xl">
            {title}
          </h1>
          {lead ? (
            <p className="max-w-3xl text-base leading-relaxed text-muted-foreground sm:text-lg">
              {lead}
            </p>
          ) : null}
        </div>
      </Container>
    </div>
  );
}
