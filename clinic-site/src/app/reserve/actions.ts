"use server";

export type ReservationType = "initial" | "follow-up" | "change-cancel";

export type ReservationFormState = {
  status: "idle" | "success" | "error";
  message: string;
};

const timeBands = ["午前", "午後", "指定なし"];

/**
 * Web予約リクエストの送信処理。
 *
 * 現時点では「送信内容をスタッフが確認し、折り返しの電話で予約を確定する」
 * リクエスト型の予約として実装している（電話予約の代替として、まず電話を減らす効果を狙う）。
 *
 * 将来的にリアルタイムの空き枠管理・自動確定を行う予約システムに置き換える場合は、
 * この関数の中身を外部予約システムのAPI呼び出しに差し替えればよい設計になっている。
 * （フォームUI側は変更不要）
 */
export async function submitReservation(
  _prevState: ReservationFormState,
  formData: FormData
): Promise<ReservationFormState> {
  const type = String(formData.get("type") ?? "") as ReservationType;
  const name = String(formData.get("name") ?? "").trim();
  const nameKana = String(formData.get("nameKana") ?? "").trim();
  const phone = String(formData.get("phone") ?? "").trim();
  const preferredDate = String(formData.get("preferredDate") ?? "").trim();
  const timeBand = String(formData.get("timeBand") ?? "").trim();
  const agree = formData.get("agree");

  if (!["initial", "follow-up", "change-cancel"].includes(type)) {
    return { status: "error", message: "ご希望の予約種別をお選びください。" };
  }
  if (!name || !nameKana) {
    return { status: "error", message: "お名前（漢字・フリガナ）をご入力ください。" };
  }
  if (!phone) {
    return { status: "error", message: "折り返しご連絡先のお電話番号をご入力ください。" };
  }
  if (type !== "change-cancel" && !preferredDate) {
    return { status: "error", message: "ご希望の日付をお選びください。" };
  }
  if (!timeBands.includes(timeBand)) {
    return { status: "error", message: "ご希望の時間帯をお選びください。" };
  }
  if (!agree) {
    return {
      status: "error",
      message: "プライバシーポリシーへの同意にチェックを入れてください。",
    };
  }

  // TODO: 実際の予約管理システム・電子カルテ連携、またはスタッフへの通知（メール/院内システム）をここに実装する。
  console.info("[reserve] reservation request received", {
    type,
    name,
    nameKana,
    phone,
    preferredDate,
    timeBand,
  });

  return {
    status: "success",
    message:
      "予約リクエストを受け付けました。内容を確認のうえ、スタッフより折り返しお電話にてご連絡し、予約を確定いたします。",
  };
}
