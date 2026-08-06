import Link from "next/link";
import { Phone, CalendarCheck } from "lucide-react";

import { desktopNav, isGroup } from "@/data/nav";
import { siteConfig } from "@/data/site";
import { Button } from "@/components/ui/button";
import { Container } from "@/components/shared/container";
import { MobileNav } from "@/components/layout/mobile-nav";
import {
  NavigationMenu,
  NavigationMenuContent,
  NavigationMenuItem,
  NavigationMenuLink,
  NavigationMenuList,
  NavigationMenuTrigger,
} from "@/components/ui/navigation-menu";

export function Header() {
  return (
    <header className="sticky top-0 z-40 border-b border-border bg-white/95 backdrop-blur supports-backdrop-blur:bg-white/80">
      <Container className="flex h-16 items-center justify-between gap-4 sm:h-20">
        <Link
          href="/"
          className="flex shrink-0 items-center gap-2 rounded text-base font-bold leading-tight whitespace-nowrap text-foreground sm:text-lg"
        >
          <span
            aria-hidden
            className="flex size-10 shrink-0 items-center justify-center rounded-full bg-brand-600 text-base font-bold text-white sm:size-11"
          >
            青
          </span>
          <span className="flex flex-col">
            <span className="hidden sm:inline lg:hidden xl:inline">
              {siteConfig.name}
            </span>
            <span className="sm:hidden lg:inline xl:hidden">
              {siteConfig.shortName}
            </span>
            <span className="hidden text-xs font-normal text-muted-foreground xl:block">
              {siteConfig.nameEn}
            </span>
          </span>
        </Link>

        <div className="hidden lg:block">
          <NavigationMenu aria-label="メインメニュー">
            <NavigationMenuList>
              {desktopNav.map((entry) =>
                isGroup(entry) ? (
                  <NavigationMenuItem key={entry.label}>
                    <NavigationMenuTrigger>{entry.label}</NavigationMenuTrigger>
                    <NavigationMenuContent>
                      <ul className="flex flex-col">
                        {entry.items.map((item) => (
                          <li key={item.href}>
                            <NavigationMenuLink asChild>
                              <Link href={item.href}>
                                {item.label}
                                {item.description ? (
                                  <span className="mt-0.5 block whitespace-normal text-xs font-normal text-muted-foreground">
                                    {item.description}
                                  </span>
                                ) : null}
                              </Link>
                            </NavigationMenuLink>
                          </li>
                        ))}
                      </ul>
                    </NavigationMenuContent>
                  </NavigationMenuItem>
                ) : (
                  <NavigationMenuItem key={entry.href}>
                    <NavigationMenuLink asChild className="px-3 py-2">
                      <Link href={entry.href}>{entry.label}</Link>
                    </NavigationMenuLink>
                  </NavigationMenuItem>
                )
              )}
            </NavigationMenuList>
          </NavigationMenu>
        </div>

        <div className="flex shrink-0 items-center gap-2">
          <Button
            asChild
            variant="secondary"
            size="sm"
            className="hidden sm:inline-flex"
          >
            <a href={siteConfig.telHref}>
              <Phone />{" "}
              <span className="hidden xl:inline">{siteConfig.telDisplay}</span>
              <span className="xl:hidden">電話する</span>
            </a>
          </Button>
          <Button asChild size="sm" className="hidden sm:inline-flex">
            <Link href="/reserve">
              <CalendarCheck /> Web予約
            </Link>
          </Button>
          <MobileNav />
        </div>
      </Container>
    </header>
  );
}
