import { faqCategories, faqItems } from "@/data/faq";
import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";

export function FaqAccordion() {
  return (
    <div className="flex flex-col gap-10">
      {faqCategories.map((category) => (
        <div key={category} className="flex flex-col gap-2">
          <h2 className="text-lg font-bold text-brand-700">{category}</h2>
          <Accordion type="multiple" className="rounded-card border border-border bg-white px-5 shadow-soft sm:px-6">
            {faqItems
              .filter((item) => item.category === category)
              .map((item) => (
                <AccordionItem key={item.question} value={item.question}>
                  <AccordionTrigger>Q. {item.question}</AccordionTrigger>
                  <AccordionContent>A. {item.answer}</AccordionContent>
                </AccordionItem>
              ))}
          </Accordion>
        </div>
      ))}
    </div>
  );
}
