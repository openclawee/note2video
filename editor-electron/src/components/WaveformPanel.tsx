import { useEffect, useImperativeHandle, useRef, forwardRef } from "react";
import WaveSurfer from "wavesurfer.js";

export type WaveformHandle = {
  play: () => void;
  pause: () => void;
  isPlaying: () => boolean;
};

type Props = {
  audioUrl: string | null;
  currentTime: number;
  onTimeUpdate: (t: number) => void;
  onReady: (duration: number) => void;
};

export const WaveformPanel = forwardRef<WaveformHandle, Props>(function WaveformPanel(
  { audioUrl, currentTime, onTimeUpdate, onReady },
  ref,
) {
  const containerRef = useRef<HTMLDivElement>(null);
  const wsRef = useRef<WaveSurfer | null>(null);
  const draggingRef = useRef(false);
  const lastSeekRef = useRef(0);

  useImperativeHandle(ref, () => ({
    play: () => {
      void wsRef.current?.play();
    },
    pause: () => {
      wsRef.current?.pause();
    },
    isPlaying: () => wsRef.current?.isPlaying() ?? false,
  }));

  useEffect(() => {
    const el = containerRef.current;
    if (!el || !audioUrl) {
      onReady(0);
      return;
    }

    const ws = WaveSurfer.create({
      container: el,
      height: 64,
      waveColor: "rgba(108, 140, 255, 0.35)",
      progressColor: "rgba(108, 140, 255, 0.95)",
      cursorColor: "#e8eaf2",
      cursorWidth: 2,
      url: audioUrl,
      dragToSeek: true,
      normalize: true,
    });

    ws.on("ready", () => {
      onReady(ws.getDuration());
    });
    ws.on("timeupdate", (t) => {
      onTimeUpdate(t);
    });
    ws.on("interaction", () => {
      draggingRef.current = true;
      onTimeUpdate(ws.getCurrentTime());
      queueMicrotask(() => {
        draggingRef.current = false;
      });
    });

    wsRef.current = ws;
    return () => {
      ws.destroy();
      wsRef.current = null;
    };
  }, [audioUrl, onReady, onTimeUpdate]);

  useEffect(() => {
    const ws = wsRef.current;
    if (!ws || !audioUrl) {
      return;
    }
    if (draggingRef.current || ws.isPlaying()) {
      return;
    }
    const now = performance.now();
    if (now - lastSeekRef.current < 40) {
      return;
    }
    const diff = Math.abs(ws.getCurrentTime() - currentTime);
    if (diff > 0.04) {
      ws.setTime(currentTime);
      lastSeekRef.current = now;
    }
  }, [audioUrl, currentTime]);

  if (!audioUrl) {
    return (
      <div className="wave-panel">
        <div className="label">旁白波形</div>
        <div
          className="wave-host"
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            color: "#6b728d",
            fontSize: 12,
          }}
        >
          未找到 audio/merged.wav（运行 voice 后可用）
        </div>
      </div>
    );
  }

  return (
    <div className="wave-panel">
      <div className="label">旁白波形（wavesurfer.js · 拖拽即 seek）</div>
      <div ref={containerRef} className="wave-host" />
    </div>
  );
});
