export type Doctor = {
  slug: string;
  role: string;
  name: string;
  nameKana: string;
  photo?: string;
  specialties: string[];
  qualifications: string[];
  career: string[];
  message: string;
  isDirector: boolean;
};

/**
 * 医師紹介データ。プレースホルダー。
 * 実際の経歴・専門・メッセージは既存サイト取得後に差し替えること。
 */
export const doctors: Doctor[] = [
  {
    slug: "director",
    role: "院長",
    name: "青山 太郎",
    nameKana: "あおやま たろう",
    specialties: ["整形外科一般", "スポーツ整形外科", "関節鏡視下手術"],
    qualifications: [
      "医学博士",
      "日本整形外科学会 整形外科専門医",
      "日本スポーツ協会公認 スポーツドクター",
    ],
    career: [
      "○○大学医学部 卒業",
      "○○大学医学部附属病院 整形外科 勤務",
      "○○病院 整形外科 医長",
      "当院 開院・院長就任",
    ],
    message:
      "「痛みを我慢しない」「早く元の生活に戻る」を大切に、お一人おひとりに合った治療をご提案します。些細なことでもお気軽にご相談ください。",
    isDirector: true,
  },
  {
    slug: "staff-doctor-1",
    role: "副院長",
    name: "青山 花子",
    nameKana: "あおやま はなこ",
    specialties: ["リハビリテーション科", "骨粗鬆症治療"],
    qualifications: ["日本リハビリテーション医学会 専門医"],
    career: ["○○大学医学部 卒業", "○○リハビリテーション病院 勤務"],
    message:
      "けがや手術のあとの生活復帰を、リハビリの立場からしっかりサポートします。",
    isDirector: false,
  },
];

export const directorDoctor = doctors.find((d) => d.isDirector) ?? doctors[0];
