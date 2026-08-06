import * as React from "react";
import { Slot } from "@radix-ui/react-slot";
import { cva, type VariantProps } from "class-variance-authority";

import { cn } from "@/lib/utils";

const buttonVariants = cva(
  "inline-flex items-center justify-center gap-2 whitespace-nowrap rounded-control text-base font-bold transition-colors disabled:pointer-events-none disabled:opacity-50 [&_svg]:pointer-events-none [&_svg]:shrink-0",
  {
    variants: {
      variant: {
        default:
          "bg-brand-600 text-white shadow-soft hover:bg-brand-700 active:bg-brand-800",
        secondary:
          "bg-white text-brand-700 ring-2 ring-inset ring-brand-200 hover:bg-brand-50",
        outline:
          "bg-transparent text-current ring-2 ring-inset ring-border hover:bg-muted",
        ghost: "bg-transparent hover:bg-muted",
        link: "text-brand-700 underline underline-offset-4 hover:text-brand-800",
        warm: "bg-accent-warm text-white shadow-soft hover:brightness-95",
      },
      size: {
        default: "h-12 px-5 [&_svg]:size-5",
        sm: "h-10 px-4 text-sm [&_svg]:size-4",
        lg: "h-14 px-7 text-lg [&_svg]:size-6",
        icon: "size-11 [&_svg]:size-5",
      },
    },
    defaultVariants: {
      variant: "default",
      size: "default",
    },
  }
);

function Button({
  className,
  variant,
  size,
  asChild = false,
  ...props
}: React.ComponentProps<"button"> &
  VariantProps<typeof buttonVariants> & { asChild?: boolean }) {
  const Comp = asChild ? Slot : "button";

  return (
    <Comp
      data-slot="button"
      className={cn(buttonVariants({ variant, size, className }))}
      {...props}
    />
  );
}

export { Button, buttonVariants };
