import { cn } from "@/lib/utils";

export function SectionHeading({
  eyebrow,
  title,
  description,
  align = "center",
  className,
}: {
  eyebrow?: string;
  title: string;
  description?: string;
  align?: "center" | "left";
  className?: string;
}) {
  return (
    <div
      className={cn(
        "flex flex-col gap-3",
        align === "center" ? "items-center text-center" : "items-start text-left",
        className
      )}
    >
      {eyebrow ? (
        <span className="text-sm font-bold tracking-wide text-brand-600">
          {eyebrow}
        </span>
      ) : null}
      <h2 className="text-2xl font-bold leading-snug text-foreground sm:text-3xl">
        {title}
      </h2>
      {description ? (
        <p
          className={cn(
            "max-w-2xl text-base leading-relaxed text-muted-foreground sm:text-lg",
            align === "center" ? "mx-auto" : ""
          )}
        >
          {description}
        </p>
      ) : null}
    </div>
  );
}
