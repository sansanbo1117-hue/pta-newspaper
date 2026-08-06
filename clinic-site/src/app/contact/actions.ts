"use server";

export type ContactFormState = {
  status: "idle" | "success" | "error";
  message: string;
};

const inquiryTypes = ["診療について", "予約について", "採用について", "その他"];

/**
 * お問い合わせフォームの送信処理。
 *
 * 現時点ではバリデーションのみ行うプロトタイプ実装。本番公開までに、
 * クリニックの受付メールアドレスへ通知するメール送信サービス（例: Resend, SendGrid など）
 * と接続すること（詳細は CONTENT-TODO.md）。
 */
export async function submitContactForm(
  _prevState: ContactFormState,
  formData: FormData
): Promise<ContactFormState> {
  const name = String(formData.get("name") ?? "").trim();
  const phone = String(formData.get("phone") ?? "").trim();
  const email = String(formData.get("email") ?? "").trim();
  const type = String(formData.get("type") ?? "").trim();
  const message = String(formData.get("message") ?? "").trim();
  const agree = formData.get("agree");

  if (!name || !message) {
    return { status: "error", message: "お名前とお問い合わせ内容をご入力ください。" };
  }
  if (!phone && !email) {
    return {
      status: "error",
      message: "折り返しご連絡できるよう、電話番号かメールアドレスのいずれかをご入力ください。",
    };
  }
  if (!inquiryTypes.includes(type)) {
    return { status: "error", message: "お問い合わせ種別をお選びください。" };
  }
  if (!agree) {
    return {
      status: "error",
      message: "プライバシーポリシーへの同意にチェックを入れてください。",
    };
  }

  // TODO: 実際のメール送信・CRM連携をここに実装する。
  console.info("[contact] inquiry received", { name, phone, email, type });

  return {
    status: "success",
    message:
      "お問い合わせを受け付けました。内容を確認のうえ、担当より折り返しご連絡いたします。",
  };
}
