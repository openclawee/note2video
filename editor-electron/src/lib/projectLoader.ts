import type { LoadedProject, ManifestJson, SubtitlesJson, TimingsJson } from "./projectTypes";

function hasApi(): boolean {
  return typeof window !== "undefined" && !!window.n2v;
}

async function readJson<T>(root: string, rel: string | undefined | null): Promise<T | null> {
  if (!rel || !hasApi()) {
    return null;
  }
  const path = await window.n2v!.joinPath(root, rel);
  try {
    const raw = await window.n2v!.readText(path);
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

async function fileExists(root: string, rel: string): Promise<boolean> {
  if (!hasApi()) {
    return false;
  }
  const path = await window.n2v!.joinPath(root, rel);
  try {
    await window.n2v!.readText(path);
    return true;
  } catch {
    return false;
  }
}

function slideTimeRangesFromTimings(
  timings: TimingsJson,
  slides: { page: number; title: string }[],
): Map<number, { start: number; end: number }> {
  const map = new Map<number, { start: number; end: number }>();
  const segs = timings.segments || [];
  for (const s of segs) {
    const page = Number(s.page);
    const start = s.start_ms / 1000;
    const end = s.end_ms / 1000;
    const cur = map.get(page);
    if (!cur) {
      map.set(page, { start, end });
    } else {
      map.set(page, {
        start: Math.min(cur.start, start),
        end: Math.max(cur.end, end),
      });
    }
  }
  for (const sl of slides) {
    if (!map.has(sl.page)) {
      map.set(sl.page, { start: 0, end: 0.3 });
    }
  }
  return map;
}

function slideTimeRangesFromManifest(slides: { page: number; duration_ms?: number }[]): Map<number, { start: number; end: number }> {
  let t = 0;
  const map = new Map<number, { start: number; end: number }>();
  for (const sl of slides) {
    const durMs = Number(sl.duration_ms || 0);
    const dur = durMs > 0 ? durMs / 1000 : 3;
    const start = t;
    const end = t + Math.max(0.25, dur);
    map.set(sl.page, { start, end });
    t = end;
  }
  return map;
}

export async function loadProject(rootDir: string): Promise<LoadedProject> {
  if (!hasApi()) {
    throw new Error("Electron API unavailable (open this app with npm run dev or electron .).");
  }
  const manifestPath = await window.n2v!.joinPath(rootDir, "manifest.json");
  const manifestRaw = await window.n2v!.readText(manifestPath);
  const manifest = JSON.parse(manifestRaw) as ManifestJson;
  const outs = manifest.outputs || {};

  const timingsRel = outs.timings || "audio/timings.json";
  const timings = await readJson<TimingsJson>(rootDir, timingsRel);

  const subRel = outs.subtitle_json || "subtitles/subtitles.json";
  const subtitles = await readJson<SubtitlesJson>(rootDir, subRel);

  const slidesRaw = manifest.slides || [];
  const slidesMeta = slidesRaw.map((s) => ({
    page: Number(s.page),
    title: String(s.title || `Slide ${s.page}`),
    duration_ms: s.duration_ms,
  }));

  const perPageTimes =
    timings && (timings.segments?.length || 0) > 0
      ? slideTimeRangesFromTimings(timings, slidesMeta)
      : slideTimeRangesFromManifest(slidesMeta);

  let total = 0;
  for (const [, range] of perPageTimes) {
    total = Math.max(total, range.end);
  }
  total = Math.max(30, total + 2);

  const slideClips = [];
  for (const sl of slidesMeta) {
    const range = perPageTimes.get(sl.page) || { start: 0, end: 3 };
    const imageRel = slidesRaw.find((x) => Number(x.page) === sl.page)?.image;
    const audioRel = slidesRaw.find((x) => Number(x.page) === sl.page)?.audio;
    let imageAbs: string | null = null;
    let audioAbs: string | null = null;
    if (imageRel) {
      imageAbs = await window.n2v!.joinPath(rootDir, imageRel);
    }
    if (audioRel) {
      const ap = await window.n2v!.joinPath(rootDir, audioRel);
      if (await fileExists(rootDir, audioRel)) {
        audioAbs = ap;
      }
    }
    slideClips.push({
      page: sl.page,
      title: sl.title,
      start: range.start,
      end: range.end,
      imageAbs,
      audioAbs,
    });
  }

  const mergedRel = "audio/merged.wav";
  let mergedAudioPath: string | null = null;
  if (await fileExists(rootDir, mergedRel)) {
    mergedAudioPath = await window.n2v!.joinPath(rootDir, mergedRel);
  }

  return {
    rootDir,
    manifest,
    manifestPath,
    timings,
    subtitles,
    mergedAudioPath,
    slideClips,
    totalDuration: total,
  };
}

export function cuesForPage(
  project: LoadedProject,
  page: number,
): Array<{ text: string; start_ms: number; end_ms: number }> {
  const subs = project.subtitles?.segments?.filter((s) => Number(s.page) === page) || [];
  if (subs.length) {
    return subs
      .map((s) => ({
        text: String(s.text || ""),
        start_ms: Number(s.start_ms),
        end_ms: Number(s.end_ms),
      }))
      .sort((a, b) => a.start_ms - b.start_ms);
  }
  const scriptSlide = project.manifest.slides?.find((s) => Number(s.page) === page);
  const script = String(scriptSlide?.script || "").trim();
  if (script) {
    return [{ text: script, start_ms: 0, end_ms: 60_000 }];
  }
  return [];
}

export function activeCue(
  cues: Array<{ text: string; start_ms: number; end_ms: number }>,
  timeSec: number,
): { text: string; index: number } | null {
  const t = timeSec * 1000;
  let best: { text: string; index: number } | null = null;
  cues.forEach((c, i) => {
    if (t >= c.start_ms && t <= c.end_ms) {
      best = { text: c.text, index: i };
    }
  });
  if (!best && cues.length) {
    return { text: cues[0].text, index: 0 };
  }
  return best;
}
