"use client";

import * as React from "react";
import Link from "next/link";
import * as Dialog from "@radix-ui/react-dialog";
import { Menu, X, Phone, CalendarCheck } from "lucide-react";

import { primaryNav } from "@/data/nav";
import { siteConfig } from "@/data/site";
import { Button } from "@/components/ui/button";

export function MobileNav() {
  const [open, setOpen] = React.useState(false);

  return (
    <Dialog.Root open={open} onOpenChange={setOpen}>
      <Dialog.Trigger asChild>
        <button
          type="button"
          className="flex size-11 items-center justify-center rounded-control text-foreground hover:bg-muted lg:hidden"
          aria-label="メニューを開く"
        >
          <Menu className="size-6" aria-hidden />
        </button>
      </Dialog.Trigger>
      <Dialog.Portal>
        <Dialog.Overlay className="fixed inset-0 z-50 bg-black/40 data-[state=open]:animate-in data-[state=open]:fade-in" />
        <Dialog.Content className="fixed inset-y-0 right-0 z-50 flex h-full w-full max-w-sm flex-col overflow-y-auto bg-white shadow-soft-lg focus:outline-none">
          <div className="flex items-center justify-between border-b border-border px-5 py-4">
            <Dialog.Title className="text-lg font-bold">メニュー</Dialog.Title>
            <Dialog.Close asChild>
              <button
                type="button"
                className="flex size-11 items-center justify-center rounded-control hover:bg-muted"
                aria-label="メニューを閉じる"
              >
                <X className="size-6" aria-hidden />
              </button>
            </Dialog.Close>
          </div>

          <nav aria-label="サイト内メニュー" className="flex-1 px-3 py-4">
            <ul className="flex flex-col">
              {primaryNav.map((item) => (
                <li key={item.href}>
                  <Link
                    href={item.href}
                    onClick={() => setOpen(false)}
                    className="block rounded-control px-3 py-4 text-lg font-bold text-foreground hover:bg-muted"
                  >
                    {item.label}
                    {item.description ? (
                      <span className="mt-0.5 block text-sm font-normal text-muted-foreground">
                        {item.description}
                      </span>
                    ) : null}
                  </Link>
                </li>
              ))}
            </ul>
          </nav>

          <div className="flex flex-col gap-3 border-t border-border p-4">
            <Button asChild size="lg" className="w-full">
              <a href={siteConfig.telHref}>
                <Phone /> {siteConfig.telDisplay} に電話する
              </a>
            </Button>
            <Button asChild size="lg" variant="secondary" className="w-full">
              <Link href="/reserve" onClick={() => setOpen(false)}>
                <CalendarCheck /> Web予約する
              </Link>
            </Button>
          </div>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}
