/**
 * サイト全体の基本情報。
 * 現時点ではプレースホルダー値。既存サイト(www.aocli.com)の内容を取得でき次第、
 * 実際の住所・電話番号・診療時間・地図座標に差し替えること（詳細は CONTENT-TODO.md）。
 */

export const siteConfig = {
  name: "青山整形外科クリニック",
  nameEn: "Aoyama Orthopedic Clinic",
  shortName: "青山整形外科",
  description:
    "青山整形外科クリニックは、地域の皆さまの「動く」を支える整形外科です。骨折・捻挫・腰痛・肩こりから、スポーツ外傷、リハビリテーションまで幅広く対応します。",
  url: "https://www.aocli.com",
  telDisplay: "○○○-○○○○-○○○○",
  telHref: "tel:0000000000",
  fax: "○○○-○○○○-○○○○",
  postalCode: "〒000-0000",
  address: {
    prefecture: "○○県",
    city: "○○市○○町",
    rest: "0-0-0",
    building: "",
  },
  get addressFull() {
    return `${this.postalCode} ${this.address.prefecture}${this.address.city}${this.address.rest}${this.address.building}`;
  },
  access: {
    train: "○○線「○○駅」より徒歩○分",
    bus: "○○バス「○○」バス停すぐ",
    parking: "専用駐車場○○台完備（no表示は現地看板をご確認ください）",
  },
  mapEmbedSrc:
    "https://www.google.com/maps?q=%E9%9D%92%E5%B1%B1%E6%95%B4%E5%BD%A2%E5%A4%96%E7%A7%91%E3%82%AF%E3%83%AA%E3%83%8B%E3%83%83%E3%82%AF&output=embed",
  mapLinkHref: "https://maps.google.com/?q=青山整形外科クリニック",
  reservationUrl: "#reserve",
  sns: {
    line: "",
    instagram: "",
  },
} as const;

export type ClinicHourRow = {
  label: string;
  am: string;
  pm: string;
  note?: string;
};

/** 診療時間。休診日・受付終了時刻などは実データ取得後に更新。 */
export const clinicHours: ClinicHourRow[] = [
  { label: "月", am: "9:00-12:30", pm: "15:00-18:30" },
  { label: "火", am: "9:00-12:30", pm: "15:00-18:30" },
  { label: "水", am: "9:00-12:30", pm: "休診" },
  { label: "木", am: "9:00-12:30", pm: "15:00-18:30" },
  { label: "金", am: "9:00-12:30", pm: "15:00-18:30" },
  { label: "土", am: "9:00-12:30", pm: "14:00-17:00" },
  { label: "日・祝", am: "休診", pm: "休診" },
];

export const closedNote = "休診日：水曜午後・日曜・祝日（詳細は「お知らせ」をご確認ください）";

export const emergencyNote =
  "急な痛み・けがなど緊急を要する場合は、まずはお電話にてご相談ください。";
