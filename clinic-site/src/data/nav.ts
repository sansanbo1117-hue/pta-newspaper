export type NavItem = {
  label: string;
  href: string;
  description?: string;
};

export const primaryNav: NavItem[] = [
  { label: "トップ", href: "/" },
  {
    label: "医院紹介",
    href: "/about",
    description: "当院の理念・特徴について",
  },
  { label: "院長紹介", href: "/director", description: "院長・医師のご紹介" },
  { label: "診療案内", href: "/services", description: "診療科目・初診の流れ" },
  { label: "診療時間", href: "/hours", description: "診療時間・休診日" },
  { label: "院内紹介", href: "/facility", description: "院内・設備のご紹介" },
  { label: "アクセス", href: "/access", description: "地図・駐車場のご案内" },
  { label: "お知らせ", href: "/news", description: "休診・お知らせ一覧" },
  { label: "よくある質問", href: "/faq", description: "予約や持ち物などのFAQ" },
  { label: "採用情報", href: "/recruit", description: "スタッフ募集" },
  { label: "お問い合わせ", href: "/contact" },
];

export const footerNav: NavItem[] = [
  ...primaryNav.filter((item) => item.href !== "/"),
  { label: "プライバシーポリシー", href: "/privacy" },
];

export const utilityActions = {
  reserve: { label: "Web予約", href: "/reserve" },
};

export type NavGroup = {
  label: string;
  items: NavItem[];
};

export type DesktopNavEntry = NavItem | NavGroup;

function isGroup(entry: DesktopNavEntry): entry is NavGroup {
  return "items" in entry;
}

function navItem(href: string): NavItem {
  const item = primaryNav.find((i) => i.href === href);
  if (!item) throw new Error(`Unknown nav href: ${href}`);
  return item;
}

/** デスクトップヘッダー用に、関連ページをグループ化した見出し数の少ないナビゲーション。 */
export const desktopNav: DesktopNavEntry[] = [
  {
    label: "医院紹介",
    items: [navItem("/about"), navItem("/director"), navItem("/facility")],
  },
  {
    label: "診療案内",
    items: [navItem("/services"), navItem("/hours")],
  },
  navItem("/access"),
  navItem("/news"),
  navItem("/faq"),
  navItem("/recruit"),
  navItem("/contact"),
];

export { isGroup };
