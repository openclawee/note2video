export type ManifestSlide = {
  page: number;
  title?: string;
  image?: string;
  script?: string;
  audio?: string;
  duration_ms?: number;
};

export type ManifestJson = {
  project_name?: string;
  input_file?: string;
  slide_count?: number;
  ratio?: string;
  resolution?: string;
  fps?: number;
  quality?: string;
  outputs?: Record<string, string>;
  slides?: ManifestSlide[];
};

export type TimingSegment = {
  page: number;
  index?: number;
  sentence_index?: number;
  text?: string;
  start_ms: number;
  end_ms: number;
  duration_ms?: number;
};

export type TimingsJson = {
  segments?: TimingSegment[];
};

export type SubtitlesJson = {
  segments?: Array<{
    index?: number;
    page: number;
    start_ms: number;
    end_ms: number;
    text: string;
  }>;
};

export type LoadedProject = {
  rootDir: string;
  manifest: ManifestJson;
  manifestPath: string;
  timings: TimingsJson | null;
  subtitles: SubtitlesJson | null;
  mergedAudioPath: string | null;
  slideClips: Array<{
    page: number;
    title: string;
    start: number;
    end: number;
    imageAbs: string | null;
    audioAbs: string | null;
  }>;
  totalDuration: number;
};
