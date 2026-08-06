export type JobPosting = {
  slug: string;
  title: string;
  employmentType: string;
  summary: string;
  requirements: string[];
  workingHours: string;
  salary: string;
  benefits: string[];
};

/** 採用情報。プレースホルダー。募集要項は実データ取得後に差し替え。 */
export const jobPostings: JobPosting[] = [
  {
    slug: "nurse",
    title: "看護師",
    employmentType: "正社員・パート",
    summary: "整形外科での外来看護、処置介助、リハビリ補助などをお任せします。",
    requirements: ["看護師免許", "経験不問（整形外科経験者歓迎）"],
    workingHours: "シフト制（応相談）",
    salary: "経験・能力に応じて優遇",
    benefits: ["社会保険完備", "交通費支給", "院内研修制度"],
  },
  {
    slug: "pt",
    title: "理学療法士",
    employmentType: "正社員",
    summary: "運動器リハビリテーションを中心に、患者さまの機能回復をサポートします。",
    requirements: ["理学療法士免許"],
    workingHours: "シフト制（応相談）",
    salary: "経験・能力に応じて優遇",
    benefits: ["社会保険完備", "交通費支給", "学会参加支援"],
  },
  {
    slug: "clerk",
    title: "医療事務",
    employmentType: "パート・アルバイト",
    summary: "受付・レセプト業務・Web予約管理などをお任せします。",
    requirements: ["医療事務経験者歓迎（未経験の方もご相談ください）"],
    workingHours: "シフト制（応相談）",
    salary: "時給応相談",
    benefits: ["交通費支給", "制服貸与"],
  },
];
