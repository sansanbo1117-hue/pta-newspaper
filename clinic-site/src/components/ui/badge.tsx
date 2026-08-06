import * as React from "react";
import { cva, type VariantProps } from "class-variance-authority";

import { cn } from "@/lib/utils";

const badgeVariants = cva(
  "inline-flex items-center rounded-full px-3 py-1 text-sm font-bold",
  {
    variants: {
      variant: {
        default: "bg-brand-50 text-brand-700 ring-1 ring-inset ring-brand-200",
        solid: "bg-brand-600 text-white",
        outline: "ring-1 ring-inset ring-border text-foreground",
        warm: "bg-orange-50 text-accent-warm ring-1 ring-inset ring-orange-200",
      },
    },
    defaultVariants: { variant: "default" },
  }
);

function Badge({
  className,
  variant,
  ...props
}: React.ComponentProps<"span"> & VariantProps<typeof badgeVariants>) {
  return (
    <span
      data-slot="badge"
      className={cn(badgeVariants({ variant, className }))}
      {...props}
    />
  );
}

export { Badge, badgeVariants };
