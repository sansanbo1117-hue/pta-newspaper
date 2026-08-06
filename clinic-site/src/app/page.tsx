import { Hero } from "@/components/home/hero";
import { QuickActions } from "@/components/home/quick-actions";
import { ServicesOverview } from "@/components/home/services-overview";
import { Features } from "@/components/home/features";
import { DirectorPreview } from "@/components/home/director-preview";
import { NewsPreview } from "@/components/home/news-preview";
import { AccessPreview } from "@/components/home/access-preview";
import { FaqPreview } from "@/components/home/faq-preview";

export default function Home() {
  return (
    <>
      <Hero />
      <QuickActions />
      <ServicesOverview />
      <Features />
      <DirectorPreview />
      <NewsPreview />
      <AccessPreview />
      <FaqPreview />
    </>
  );
}
