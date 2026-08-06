import type { LucideIcon } from "lucide-react";
import {
  Bone,
  Activity,
  HeartPulse,
  Dumbbell,
  Baby,
  ShieldPlus,
  Stethoscope,
  Waves,
} from "lucide-react";

export type Service = {
  slug: string;
  title: string;
  icon: LucideIcon;
  summary: string;
  details: string[];
};

/** 診療案内。プレースホルダー内容。実際の診療科目・詳細文は既存サイト取得後に差し替え。 */
export const services: Service[] = [
  {
    slug: "general",
    title: "整形外科一般",
    icon: Bone,
    summary: "骨折・脱臼・捻挫・打撲など、けが全般に対応します。",
    details: [
      "骨折・脱臼・捻挫・打撲の診断と治療",
      "レントゲン検査による迅速な診断",
      "ギプス固定・装具療法",
    ],
  },
  {
    slug: "sports",
    title: "スポーツ整形外科",
    icon: Dumbbell,
    summary: "スポーツによるけがの治療から競技復帰までサポートします。",
    details: [
      "靭帯損傷・半月板損傷などのスポーツ外傷",
      "疲労骨折・使いすぎ症候群（オーバーユース）",
      "競技復帰に向けたリハビリ指導",
    ],
  },
  {
    slug: "spine",
    title: "腰痛・肩こり外来",
    icon: Activity,
    summary: "慢性的な腰痛・肩こり・首の痛みを丁寧に診察します。",
    details: ["腰椎椎間板ヘルニア・脊柱管狭窄症", "頚椎症・寝違え", "姿勢・生活指導"],
  },
  {
    slug: "rehabilitation",
    title: "リハビリテーション",
    icon: HeartPulse,
    summary: "理学療法士による個別リハビリで機能回復を支援します。",
    details: [
      "運動器リハビリテーション",
      "術後・骨折後の機能回復訓練",
      "物理療法（温熱・電気治療など）",
    ],
  },
  {
    slug: "osteoporosis",
    title: "骨粗鬆症検診・治療",
    icon: ShieldPlus,
    summary: "骨密度測定による早期発見と、その方に合った治療を行います。",
    details: ["骨密度測定（DXA法）", "薬物療法・食事/運動指導", "定期フォローアップ"],
  },
  {
    slug: "pediatric",
    title: "小児整形外科",
    icon: Baby,
    summary: "お子さまの成長期特有のけが・症状にも対応します。",
    details: ["成長痛・オスグッド病", "運動器検診のご相談", "先天性股関節疾患のフォロー"],
  },
  {
    slug: "injection",
    title: "関節注射・物理療法",
    icon: Stethoscope,
    summary: "症状に応じた注射治療や物理療法をご用意しています。",
    details: ["ヒアルロン酸注射", "神経ブロック注射", "超音波治療・牽引療法"],
  },
  {
    slug: "labor",
    title: "労災・交通事故対応",
    icon: Waves,
    summary: "労災保険・自賠責保険を用いた治療にも対応しています。",
    details: ["労災指定医療機関", "交通事故によるけがの診断書作成", "保険会社との連携"],
  },
];

export type FirstVisitStep = {
  step: number;
  title: string;
  description: string;
};

export const firstVisitSteps: FirstVisitStep[] = [
  {
    step: 1,
    title: "受付",
    description:
      "保険証（と、あればお薬手帳・紹介状）をご提示ください。Web予約の方は予約番号をお伝えいただくとスムーズです。",
  },
  {
    step: 2,
    title: "問診",
    description:
      "症状やお困りごとについて、問診票のご記入とスタッフによる聞き取りを行います。",
  },
  {
    step: 3,
    title: "診察・検査",
    description:
      "医師が診察し、必要に応じてレントゲンなどの検査を行い、症状を確認します。",
  },
  {
    step: 4,
    title: "説明・治療",
    description: "検査結果と診断について分かりやすくご説明し、治療方針をご提案します。",
  },
  {
    step: 5,
    title: "会計",
    description: "お会計後、次回のご予約が必要な場合はご案内します。",
  },
];
