"use client";

import * as React from "react";
import * as NavigationMenuPrimitive from "@radix-ui/react-navigation-menu";
import { ChevronDown } from "lucide-react";

import { cn } from "@/lib/utils";

function NavigationMenu({
  className,
  children,
  ...props
}: React.ComponentProps<typeof NavigationMenuPrimitive.Root>) {
  return (
    <NavigationMenuPrimitive.Root
      data-slot="navigation-menu"
      className={cn("relative flex items-center", className)}
      {...props}
    >
      {children}
      <NavigationMenuViewport />
    </NavigationMenuPrimitive.Root>
  );
}

function NavigationMenuList({
  className,
  ...props
}: React.ComponentProps<typeof NavigationMenuPrimitive.List>) {
  return (
    <NavigationMenuPrimitive.List
      data-slot="navigation-menu-list"
      className={cn("flex items-center gap-1", className)}
      {...props}
    />
  );
}

const NavigationMenuItem = NavigationMenuPrimitive.Item;

function NavigationMenuTrigger({
  className,
  children,
  ...props
}: React.ComponentProps<typeof NavigationMenuPrimitive.Trigger>) {
  return (
    <NavigationMenuPrimitive.Trigger
      data-slot="navigation-menu-trigger"
      className={cn(
        "group flex shrink-0 items-center gap-1 whitespace-nowrap rounded-control px-2.5 py-2 text-sm font-bold text-foreground/90 outline-none hover:bg-muted hover:text-brand-700 data-[state=open]:bg-muted data-[state=open]:text-brand-700 xl:px-3",
        className
      )}
      {...props}
    >
      {children}
      <ChevronDown
        className="size-3.5 transition-transform duration-200 group-data-[state=open]:rotate-180"
        aria-hidden
      />
    </NavigationMenuPrimitive.Trigger>
  );
}

function NavigationMenuContent({
  className,
  ...props
}: React.ComponentProps<typeof NavigationMenuPrimitive.Content>) {
  return (
    <NavigationMenuPrimitive.Content
      data-slot="navigation-menu-content"
      className={cn(
        "min-w-64 rounded-card border border-border bg-white p-2 shadow-soft-lg data-[motion=from-end]:animate-in data-[motion=from-start]:animate-in data-[motion=to-end]:animate-out data-[motion=to-start]:animate-out data-[motion=from-end]:slide-in-from-right-4 data-[motion=from-start]:slide-in-from-left-4 data-[motion=to-end]:slide-out-to-right-4 data-[motion=to-start]:slide-out-to-left-4",
        className
      )}
      {...props}
    />
  );
}

function NavigationMenuLink({
  className,
  ...props
}: React.ComponentProps<typeof NavigationMenuPrimitive.Link>) {
  return (
    <NavigationMenuPrimitive.Link
      data-slot="navigation-menu-link"
      className={cn(
        "block shrink-0 whitespace-nowrap rounded-control px-3 py-2.5 text-sm font-bold text-foreground hover:bg-muted hover:text-brand-700",
        className
      )}
      {...props}
    />
  );
}

function NavigationMenuViewport() {
  return (
    <div className="absolute top-full left-0 z-40 flex justify-center">
      <NavigationMenuPrimitive.Viewport className="mt-2 origin-top overflow-hidden rounded-card bg-white transition-[width,height] duration-200 data-[state=closed]:animate-out data-[state=closed]:fade-out data-[state=open]:animate-in data-[state=open]:fade-in" />
    </div>
  );
}

export {
  NavigationMenu,
  NavigationMenuList,
  NavigationMenuItem,
  NavigationMenuTrigger,
  NavigationMenuContent,
  NavigationMenuLink,
};
