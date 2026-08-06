import { Car, DoorOpen, MapPin, TrainFront } from "lucide-react";

import { siteConfig } from "@/data/site";

export function AccessMap({ showPhotos = true }: { showPhotos?: boolean }) {
  return (
    <div className="flex flex-col gap-6">
      <div className="overflow-hidden rounded-card shadow-soft">
        <iframe
          title={`${siteConfig.name} 地図`}
          src={siteConfig.mapEmbedSrc}
          className="h-80 w-full border-0 sm:h-[420px]"
          loading="lazy"
          referrerPolicy="no-referrer-when-downgrade"
        />
      </div>

      <ul className="grid gap-4 sm:grid-cols-3">
        <li className="flex items-start gap-3 rounded-card bg-white p-5 shadow-soft">
          <MapPin className="mt-0.5 size-5 shrink-0 text-brand-600" aria-hidden />
          <div>
            <p className="font-bold">住所</p>
            <p className="text-sm leading-relaxed text-muted-foreground">
              {siteConfig.addressFull}
            </p>
          </div>
        </li>
        <li className="flex items-start gap-3 rounded-card bg-white p-5 shadow-soft">
          <TrainFront className="mt-0.5 size-5 shrink-0 text-brand-600" aria-hidden />
          <div>
            <p className="font-bold">電車・バス</p>
            <p className="text-sm leading-relaxed text-muted-foreground">
              {siteConfig.access.train}
              <br />
              {siteConfig.access.bus}
            </p>
          </div>
        </li>
        <li className="flex items-start gap-3 rounded-card bg-white p-5 shadow-soft">
          <Car className="mt-0.5 size-5 shrink-0 text-brand-600" aria-hidden />
          <div>
            <p className="font-bold">お車でお越しの方</p>
            <p className="text-sm leading-relaxed text-muted-foreground">
              {siteConfig.access.parking}
            </p>
          </div>
        </li>
      </ul>

      {showPhotos ? (
        <div className="grid gap-4 sm:grid-cols-2">
          <figure className="overflow-hidden rounded-card bg-brand-50 shadow-soft">
            <div className="flex aspect-video flex-col items-center justify-center gap-2 text-brand-700">
              <Car className="size-10" aria-hidden />
              <p className="text-sm font-bold">駐車場の写真（準備中）</p>
            </div>
            <figcaption className="px-4 py-3 text-sm text-muted-foreground">
              駐車場から入口までの動線が分かる写真に差し替え予定です。
            </figcaption>
          </figure>
          <figure className="overflow-hidden rounded-card bg-brand-50 shadow-soft">
            <div className="flex aspect-video flex-col items-center justify-center gap-2 text-brand-700">
              <DoorOpen className="size-10" aria-hidden />
              <p className="text-sm font-bold">入口の写真（準備中）</p>
            </div>
            <figcaption className="px-4 py-3 text-sm text-muted-foreground">
              初診の方が迷わないよう、入口の見た目が分かる写真に差し替え予定です。
            </figcaption>
          </figure>
        </div>
      ) : null}
    </div>
  );
}
