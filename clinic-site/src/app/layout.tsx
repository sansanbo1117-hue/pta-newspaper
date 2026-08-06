import type { Metadata, Viewport } from "next";
import { Noto_Sans_JP } from "next/font/google";

import "./globals.css";
import { siteConfig } from "@/data/site";
import { Header } from "@/components/layout/header";
import { Footer } from "@/components/layout/footer";
import { MobileActionBar } from "@/components/layout/mobile-action-bar";

const notoSansJP = Noto_Sans_JP({
  variable: "--font-noto-sans-jp",
  subsets: ["latin"],
  weight: ["400", "500", "700", "900"],
  display: "swap",
});

export const metadata: Metadata = {
  metadataBase: new URL(siteConfig.url),
  title: {
    default: `${siteConfig.name}｜公式サイト・患者ポータル`,
    template: `%s｜${siteConfig.name}`,
  },
  description: siteConfig.description,
  keywords: [
    "整形外科",
    "リハビリテーション",
    "スポーツ整形外科",
    siteConfig.name,
    siteConfig.address.city,
  ],
  openGraph: {
    type: "website",
    locale: "ja_JP",
    siteName: siteConfig.name,
    title: siteConfig.name,
    description: siteConfig.description,
    url: siteConfig.url,
  },
  twitter: {
    card: "summary_large_image",
    title: siteConfig.name,
    description: siteConfig.description,
  },
  // 住所・電話番号・本文の多くがプレースホルダーのため、実データ反映まではnoindexにしておく。
  // 公開時は CONTENT-TODO.md の手順に従い index: true / follow: true に戻すこと。
  robots: {
    index: false,
    follow: false,
  },
  alternates: {
    canonical: "/",
  },
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  themeColor: "#1c63c4",
};

const clinicJsonLd = {
  "@context": "https://schema.org",
  "@type": "MedicalClinic",
  name: siteConfig.name,
  alternateName: siteConfig.nameEn,
  url: siteConfig.url,
  telephone: siteConfig.telDisplay,
  address: {
    "@type": "PostalAddress",
    streetAddress: `${siteConfig.address.city}${siteConfig.address.rest}`,
    addressRegion: siteConfig.address.prefecture,
    postalCode: siteConfig.postalCode.replace("〒", ""),
    addressCountry: "JP",
  },
  medicalSpecialty: "Orthopedic",
  priceRange: "$$",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ja" data-scroll-behavior="smooth" className={notoSansJP.variable}>
      <body className="flex min-h-dvh flex-col bg-background text-foreground antialiased">
        <a href="#main-content" className="skip-link">
          本文へスキップ
        </a>
        <script
          type="application/ld+json"
          dangerouslySetInnerHTML={{ __html: JSON.stringify(clinicJsonLd) }}
        />
        <Header />
        <main id="main-content" className="flex-1 pb-16 lg:pb-0">
          {children}
        </main>
        <Footer />
        <MobileActionBar />
      </body>
    </html>
  );
}
