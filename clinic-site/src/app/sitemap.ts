import type { MetadataRoute } from "next";

import { siteConfig } from "@/data/site";
import { newsItems } from "@/data/news";

const staticPaths = [
  "",
  "/about",
  "/director",
  "/services",
  "/hours",
  "/access",
  "/facility",
  "/news",
  "/recruit",
  "/reserve",
  "/faq",
  "/contact",
  "/privacy",
];

export default function sitemap(): MetadataRoute.Sitemap {
  const staticEntries: MetadataRoute.Sitemap = staticPaths.map((path) => ({
    url: `${siteConfig.url}${path}`,
    changeFrequency: path === "" ? "weekly" : "monthly",
    priority: path === "" ? 1 : 0.6,
  }));

  const newsEntries: MetadataRoute.Sitemap = newsItems.map((item) => ({
    url: `${siteConfig.url}/news/${item.slug}`,
    lastModified: item.date,
    changeFrequency: "never",
    priority: 0.4,
  }));

  return [...staticEntries, ...newsEntries];
}
