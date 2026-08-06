"use client";

import { useActionState } from "react";
import { CheckCircle2, CircleAlert, Send } from "lucide-react";

import { submitContactForm, type ContactFormState } from "@/app/contact/actions";
import { Button } from "@/components/ui/button";

const initialState: ContactFormState = { status: "idle", message: "" };

const inquiryTypes = ["診療について", "予約について", "採用について", "その他"];

function Field({
  label,
  htmlFor,
  required,
  children,
}: {
  label: string;
  htmlFor: string;
  required?: boolean;
  children: React.ReactNode;
}) {
  return (
    <div className="flex flex-col gap-2">
      <label htmlFor={htmlFor} className="font-bold">
        {label}
        {required ? <span className="ml-1 text-rose-600">*</span> : null}
      </label>
      {children}
    </div>
  );
}

const inputClass =
  "h-12 w-full rounded-control border border-border bg-white px-4 text-base outline-none focus-visible:border-brand-500";

export function ContactForm() {
  const [state, formAction, isPending] = useActionState(
    submitContactForm,
    initialState
  );

  if (state.status === "success") {
    return (
      <div
        role="status"
        className="flex flex-col items-center gap-3 rounded-card bg-emerald-50 p-8 text-center text-emerald-800 ring-1 ring-inset ring-emerald-200"
      >
        <CheckCircle2 className="size-10" aria-hidden />
        <p className="text-lg font-bold">送信が完了しました</p>
        <p className="text-sm leading-relaxed">{state.message}</p>
      </div>
    );
  }

  return (
    <form action={formAction} className="flex flex-col gap-6" noValidate>
      {state.status === "error" ? (
        <div
          role="alert"
          className="flex items-start gap-2 rounded-control bg-rose-50 p-4 text-sm text-rose-700 ring-1 ring-inset ring-rose-200"
        >
          <CircleAlert className="mt-0.5 size-4 shrink-0" aria-hidden />
          {state.message}
        </div>
      ) : null}

      <Field label="お名前" htmlFor="name" required>
        <input id="name" name="name" required autoComplete="name" className={inputClass} />
      </Field>

      <div className="grid gap-6 sm:grid-cols-2">
        <Field label="電話番号" htmlFor="phone">
          <input
            id="phone"
            name="phone"
            type="tel"
            autoComplete="tel"
            className={inputClass}
            placeholder="000-0000-0000"
          />
        </Field>
        <Field label="メールアドレス" htmlFor="email">
          <input
            id="email"
            name="email"
            type="email"
            autoComplete="email"
            className={inputClass}
            placeholder="example@mail.com"
          />
        </Field>
      </div>

      <Field label="お問い合わせ種別" htmlFor="type" required>
        <select id="type" name="type" required defaultValue="" className={inputClass}>
          <option value="" disabled>
            選択してください
          </option>
          {inquiryTypes.map((t) => (
            <option key={t} value={t}>
              {t}
            </option>
          ))}
        </select>
      </Field>

      <Field label="お問い合わせ内容" htmlFor="message" required>
        <textarea
          id="message"
          name="message"
          required
          rows={6}
          className="w-full rounded-control border border-border bg-white px-4 py-3 text-base outline-none focus-visible:border-brand-500"
        />
      </Field>

      <label htmlFor="agree" className="flex items-start gap-3 text-sm text-muted-foreground">
        <input
          id="agree"
          name="agree"
          type="checkbox"
          required
          className="mt-1 size-5 shrink-0 rounded border-border text-brand-600 focus-visible:outline-brand-500"
        />
        <a href="/privacy" className="underline hover:text-brand-700">
          プライバシーポリシー
        </a>
        の内容に同意のうえ送信します。
      </label>

      <Button type="submit" size="lg" disabled={isPending} className="w-full sm:w-fit">
        <Send /> {isPending ? "送信中..." : "送信する"}
      </Button>
    </form>
  );
}
