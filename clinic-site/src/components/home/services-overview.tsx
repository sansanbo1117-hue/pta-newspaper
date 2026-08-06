import Link from "next/link";
import { ArrowRight } from "lucide-react";

import { services } from "@/data/services";
import { Container } from "@/components/shared/container";
import { SectionHeading } from "@/components/shared/section-heading";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";

export function ServicesOverview() {
  return (
    <section className="py-14 sm:py-20">
      <Container className="flex flex-col gap-10">
        <SectionHeading
          eyebrow="SERVICES"
          title="診療案内"
          description="骨折・スポーツ外傷から慢性的な腰痛・肩こり、リハビリテーションまで幅広く対応しています。"
        />

        <ul className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          {services.slice(0, 8).map((service) => {
            const Icon = service.icon;
            return (
              <li key={service.slug}>
                <Card className="h-full transition-shadow hover:shadow-soft-lg">
                  <CardHeader>
                    <span className="mb-2 flex size-12 items-center justify-center rounded-full bg-brand-50 text-brand-700">
                      <Icon className="size-6" aria-hidden />
                    </span>
                    <CardTitle>{service.title}</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <p className="text-sm leading-relaxed text-muted-foreground">
                      {service.summary}
                    </p>
                  </CardContent>
                </Card>
              </li>
            );
          })}
        </ul>

        <div className="flex justify-center">
          <Link
            href="/services"
            className="inline-flex items-center gap-1.5 text-base font-bold text-brand-700 hover:underline"
          >
            診療案内をもっと見る
            <ArrowRight className="size-4" aria-hidden />
          </Link>
        </div>
      </Container>
    </section>
  );
}
