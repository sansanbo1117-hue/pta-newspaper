import Link from "next/link";
import { MapPin, Phone, Printer, Clock } from "lucide-react";

import { footerNav } from "@/data/nav";
import { siteConfig, clinicHours } from "@/data/site";
import { Container } from "@/components/shared/container";

export function Footer() {
  return (
    <footer className="border-t border-border bg-muted">
      <Container className="grid gap-10 py-12 sm:py-16 md:grid-cols-[1.2fr_1fr_1fr]">
        <div className="flex flex-col gap-4">
          <div className="flex items-center gap-2">
            <span
              aria-hidden
              className="flex size-10 items-center justify-center rounded-full bg-brand-600 text-base font-bold text-white"
            >
              青
            </span>
            <span className="text-lg font-bold">{siteConfig.name}</span>
          </div>
          <p className="max-w-sm text-sm leading-relaxed text-muted-foreground">
            {siteConfig.description}
          </p>
          <ul className="flex flex-col gap-2 text-sm text-muted-foreground">
            <li className="flex items-start gap-2">
              <MapPin className="mt-0.5 size-4 shrink-0 text-brand-600" aria-hidden />
              {siteConfig.addressFull}
            </li>
            <li className="flex items-center gap-2">
              <Phone className="size-4 shrink-0 text-brand-600" aria-hidden />
              <a href={siteConfig.telHref} className="hover:text-brand-700">
                {siteConfig.telDisplay}
              </a>
            </li>
            <li className="flex items-center gap-2">
              <Printer className="size-4 shrink-0 text-brand-600" aria-hidden />
              FAX：{siteConfig.fax}
            </li>
          </ul>
        </div>

        <div>
          <h2 className="mb-4 text-sm font-bold text-foreground">サイトメニュー</h2>
          <ul className="grid grid-cols-1 gap-2.5 text-sm text-muted-foreground sm:grid-cols-2">
            {footerNav.map((item) => (
              <li key={item.href}>
                <Link href={item.href} className="hover:text-brand-700">
                  {item.label}
                </Link>
              </li>
            ))}
          </ul>
        </div>

        <div>
          <h2 className="mb-4 flex items-center gap-2 text-sm font-bold text-foreground">
            <Clock className="size-4 text-brand-600" aria-hidden />
            診療時間
          </h2>
          <table className="w-full border-collapse text-sm text-muted-foreground">
            <tbody>
              {clinicHours.map((row) => (
                <tr key={row.label} className="border-b border-border/70">
                  <th scope="row" className="py-1.5 pr-2 text-left font-bold text-foreground">
                    {row.label}
                  </th>
                  <td className="py-1.5 pr-2">{row.am}</td>
                  <td className="py-1.5">{row.pm}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <Link
            href="/hours"
            className="mt-3 inline-block text-sm font-bold text-brand-700 hover:underline"
          >
            診療時間の詳細を見る →
          </Link>
        </div>
      </Container>

      <div className="border-t border-border py-6">
        <Container className="flex flex-col items-center justify-between gap-3 text-xs text-muted-foreground sm:flex-row">
          <p>
            &copy; {new Date().getFullYear()} {siteConfig.name}
          </p>
          <Link href="/privacy" className="hover:text-brand-700">
            プライバシーポリシー
          </Link>
        </Container>
      </div>
    </footer>
  );
}
