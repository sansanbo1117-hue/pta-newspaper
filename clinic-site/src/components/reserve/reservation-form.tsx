"use client";

import { useActionState, useId, useState } from "react";
import { CheckCircle2, CircleAlert, Send } from "lucide-react";

import {
  submitReservation,
  type ReservationFormState,
  type ReservationType,
} from "@/app/reserve/actions";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

const initialState: ReservationFormState = { status: "idle", message: "" };

const typeOptions: { value: ReservationType; label: string; description: string }[] = [
  { value: "initial", label: "初診予約", description: "はじめて受診される方" },
  { value: "follow-up", label: "再診予約", description: "すでに通院中の方" },
  {
    value: "change-cancel",
    label: "予約変更・キャンセル",
    description: "既存のご予約を変更・取消したい方",
  },
];

const timeBands = ["午前", "午後", "指定なし"];

const inputClass =
  "h-12 w-full rounded-control border border-border bg-white px-4 text-base outline-none focus-visible:border-brand-500";

export function ReservationForm() {
  const [state, formAction, isPending] = useActionState(
    submitReservation,
    initialState
  );
  const [type, setType] = useState<ReservationType>("initial");
  const groupId = useId();

  if (state.status === "success") {
    return (
      <div
        role="status"
        className="flex flex-col items-center gap-3 rounded-card bg-emerald-50 p-8 text-center text-emerald-800 ring-1 ring-inset ring-emerald-200"
      >
        <CheckCircle2 className="size-10" aria-hidden />
        <p className="text-lg font-bold">予約リクエストを受け付けました</p>
        <p className="text-sm leading-relaxed">{state.message}</p>
      </div>
    );
  }

  return (
    <form action={formAction} className="flex flex-col gap-8" noValidate>
      <fieldset className="flex flex-col gap-3">
        <legend className="mb-1 font-bold">ご希望の予約種別</legend>
        <div
          role="radiogroup"
          aria-labelledby={groupId}
          className="grid gap-3 sm:grid-cols-3"
        >
          {typeOptions.map((opt) => {
            const checked = type === opt.value;
            return (
              <label
                key={opt.value}
                className={cn(
                  "flex cursor-pointer flex-col gap-1 rounded-card border-2 p-4 transition-colors",
                  checked
                    ? "border-brand-600 bg-brand-50"
                    : "border-border bg-white hover:border-brand-200"
                )}
              >
                <input
                  type="radio"
                  name="type"
                  value={opt.value}
                  checked={checked}
                  onChange={() => setType(opt.value)}
                  className="sr-only"
                />
                <span className="font-bold">{opt.label}</span>
                <span className="text-xs text-muted-foreground">
                  {opt.description}
                </span>
              </label>
            );
          })}
        </div>
      </fieldset>

      {state.status === "error" ? (
        <div
          role="alert"
          className="flex items-start gap-2 rounded-control bg-rose-50 p-4 text-sm text-rose-700 ring-1 ring-inset ring-rose-200"
        >
          <CircleAlert className="mt-0.5 size-4 shrink-0" aria-hidden />
          {state.message}
        </div>
      ) : null}

      <div className="grid gap-6 sm:grid-cols-2">
        <div className="flex flex-col gap-2">
          <label htmlFor="name" className="font-bold">
            お名前<span className="ml-1 text-rose-600">*</span>
          </label>
          <input id="name" name="name" required autoComplete="name" className={inputClass} />
        </div>
        <div className="flex flex-col gap-2">
          <label htmlFor="nameKana" className="font-bold">
            フリガナ<span className="ml-1 text-rose-600">*</span>
          </label>
          <input id="nameKana" name="nameKana" required className={inputClass} />
        </div>
      </div>

      <div className="grid gap-6 sm:grid-cols-2">
        <div className="flex flex-col gap-2">
          <label htmlFor="phone" className="font-bold">
            電話番号<span className="ml-1 text-rose-600">*</span>
          </label>
          <input
            id="phone"
            name="phone"
            type="tel"
            required
            autoComplete="tel"
            placeholder="000-0000-0000"
            className={inputClass}
          />
        </div>
        {type !== "change-cancel" ? (
          <div className="flex flex-col gap-2">
            <label htmlFor="preferredDate" className="font-bold">
              ご希望日<span className="ml-1 text-rose-600">*</span>
            </label>
            <input
              id="preferredDate"
              name="preferredDate"
              type="date"
              required
              className={inputClass}
            />
          </div>
        ) : (
          <div className="flex flex-col gap-2">
            <label htmlFor="preferredDate" className="font-bold">
              現在のご予約日（分かる範囲で）
            </label>
            <input id="preferredDate" name="preferredDate" type="date" className={inputClass} />
          </div>
        )}
      </div>

      <fieldset className="flex flex-col gap-2">
        <legend className="font-bold">
          ご希望の時間帯<span className="ml-1 text-rose-600">*</span>
        </legend>
        <div className="flex flex-wrap gap-3">
          {timeBands.map((band) => (
            <label
              key={band}
              className="flex items-center gap-2 rounded-control border border-border px-4 py-2.5 has-[:checked]:border-brand-600 has-[:checked]:bg-brand-50"
            >
              <input type="radio" name="timeBand" value={band} required className="size-4" />
              {band}
            </label>
          ))}
        </div>
      </fieldset>

      <div className="flex flex-col gap-2">
        <label htmlFor="notes" className="font-bold">
          症状・ご要望など（任意）
        </label>
        <textarea
          id="notes"
          name="notes"
          rows={4}
          className="w-full rounded-control border border-border bg-white px-4 py-3 text-base outline-none focus-visible:border-brand-500"
          placeholder="例：3日前から右ひざが痛い／〇月〇日の予約を〇月〇日に変更したい　など"
        />
      </div>

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
        <Send /> {isPending ? "送信中..." : "予約をリクエストする"}
      </Button>
    </form>
  );
}
