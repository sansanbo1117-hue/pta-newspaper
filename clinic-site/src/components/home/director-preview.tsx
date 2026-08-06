import Link from "next/link";
import { ArrowRight, Stethoscope } from "lucide-react";

import { directorDoctor } from "@/data/doctors";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";
import { Button } from "@/components/ui/button";

export function DirectorPreview() {
  return (
    <section className="py-14 sm:py-20">
      <Container className="flex flex-col gap-10">
        <SectionHeading eyebrow="DIRECTOR" title="院長紹介" />

        <div className="grid gap-8 rounded-card bg-white p-6 shadow-soft sm:p-10 lg:grid-cols-[220px_1fr] lg:items-center">
          <div
            role="img"
            aria-label={`${directorDoctor.name}院長の写真（準備中）`}
            className="mx-auto flex size-40 items-center justify-center rounded-full bg-brand-50 text-brand-700 lg:size-52"
          >
            <Stethoscope className="size-16" aria-hidden />
          </div>
          <div className="flex flex-col gap-4 text-center lg:text-left">
            <div>
              <p className="text-sm font-bold text-brand-600">
                {directorDoctor.role}
              </p>
              <p className="text-2xl font-bold">
                {directorDoctor.name}
                <span className="ml-2 text-base font-normal text-muted-foreground">
                  {directorDoctor.nameKana}
                </span>
              </p>
            </div>
            <p className="leading-relaxed text-muted-foreground">
              {directorDoctor.message}
            </p>
            <div className="flex justify-center lg:justify-start">
              <Button asChild variant="secondary">
                <Link href="/director">
                  院長・医師紹介をもっと見る <ArrowRight className="size-4" />
                </Link>
              </Button>
            </div>
          </div>
        </div>
      </Container>
    </section>
  );
}
