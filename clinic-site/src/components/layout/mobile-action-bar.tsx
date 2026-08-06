import Link from "next/link";
import { Phone, CalendarCheck, MapPin, Clock } from "lucide-react";

import { siteConfig } from "@/data/site";

const actions = [
  { href: "/access", label: "アクセス", icon: MapPin },
  { href: "/hours", label: "診療時間", icon: Clock },
] as const;

export function MobileActionBar() {
  return (
    <div
      className="fixed inset-x-0 bottom-0 z-40 border-t border-border bg-white/95 backdrop-blur pb-[env(safe-area-inset-bottom)] lg:hidden"
      role="navigation"
      aria-label="クイックアクション"
    >
      <ul className="grid grid-cols-4">
        <li>
          <a
            href={siteConfig.telHref}
            className="flex h-16 flex-col items-center justify-center gap-0.5 bg-brand-600 text-white"
          >
            <Phone className="size-5" aria-hidden />
            <span className="text-xs font-bold">電話する</span>
          </a>
        </li>
        <li>
          <Link
            href="/reserve"
            className="flex h-16 flex-col items-center justify-center gap-0.5 bg-accent-warm text-white"
          >
            <CalendarCheck className="size-5" aria-hidden />
            <span className="text-xs font-bold">Web予約</span>
          </Link>
        </li>
        {actions.map(({ href, label, icon: Icon }) => (
          <li key={href}>
            <Link
              href={href}
              className="flex h-16 flex-col items-center justify-center gap-0.5 text-foreground hover:bg-muted"
            >
              <Icon className="size-5 text-brand-600" aria-hidden />
              <span className="text-xs font-bold">{label}</span>
            </Link>
          </li>
        ))}
      </ul>
    </div>
  );
}
