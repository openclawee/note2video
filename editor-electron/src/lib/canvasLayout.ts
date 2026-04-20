/** Mirrors `note2video.video_canvas` so preview matches pipeline output dimensions. */

export function normalizeRatio(value: string | undefined): "16:9" | "9:16" | "1:1" {
  const raw = String(value || "")
    .trim()
    .toLowerCase()
    .replace(/：/g, ":")
    .replace(/x/g, ":")
    .replace(/\s/g, "");
  if (raw === "16:9" || raw === "9:16" || raw === "1:1") {
    return raw;
  }
  return "16:9";
}

export function normalizeResolution(value: string | undefined): "720p" | "1080p" | "1440p" {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "720p" || raw === "1080p" || raw === "1440p") {
    return raw;
  }
  return "1080p";
}

function ratioBaseSize(ratio: string): { w: number; h: number } {
  if (ratio === "16:9") {
    return { w: 1920, h: 1080 };
  }
  if (ratio === "9:16") {
    return { w: 1080, h: 1920 };
  }
  return { w: 1080, h: 1080 };
}

export function canvasSize(
  ratio: string | undefined,
  resolution: string | undefined,
): { w: number; h: number } {
  const r = normalizeRatio(ratio);
  const base = ratioBaseSize(r);
  const normalized = normalizeResolution(resolution);
  const scaleMap: Record<string, number> = {
    "720p": 2 / 3,
    "1080p": 1,
    "1440p": 4 / 3,
  };
  const scale = scaleMap[normalized] ?? 1;
  return {
    w: Math.round(base.w * scale),
    h: Math.round(base.h * scale),
  };
}
