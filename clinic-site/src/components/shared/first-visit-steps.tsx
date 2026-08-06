import { ArrowRight, ClipboardList, CreditCard, Stethoscope, UserCheck, Microscope } from "lucide-react";

import { firstVisitSteps } from "@/data/services";

const icons = [UserCheck, ClipboardList, Microscope, Stethoscope, CreditCard];

export function FirstVisitSteps() {
  return (
    <ol className="grid gap-4 sm:grid-cols-5">
      {firstVisitSteps.map((step, index) => {
        const Icon = icons[index];
        return (
          <li key={step.step} className="relative flex">
            <div className="flex w-full flex-col items-center gap-3 rounded-card bg-white p-5 text-center shadow-soft">
              <span className="flex size-12 items-center justify-center rounded-full bg-brand-600 text-lg font-bold text-white">
                {step.step}
              </span>
              <Icon className="size-7 text-brand-600" aria-hidden />
              <h3 className="font-bold">{step.title}</h3>
              <p className="text-sm leading-relaxed text-muted-foreground">
                {step.description}
              </p>
            </div>
            {index < firstVisitSteps.length - 1 ? (
              <ArrowRight
                aria-hidden
                className="absolute top-1/2 -right-3 z-10 hidden size-6 -translate-y-1/2 text-brand-300 sm:block"
              />
            ) : null}
          </li>
        );
      })}
    </ol>
  );
}
