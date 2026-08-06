import type { MetadataRoute } from "next";

import { siteConfig } from "@/data/site";

// プレースホルダーの住所・電話番号・本文で検索エンジンに登録されないよう、
// 実データ反映・公開判断が済むまでは全面的にクロールを禁止する。
// 公開時は CONTENT-TODO.md の手順に従い allow: "/" に戻すこと。
export default function robots(): MetadataRoute.Robots {
  return {
    rules: {
      userAgent: "*",
      disallow: "/",
    },
    sitemap: `${siteConfig.url}/sitemap.xml`,
  };
}
