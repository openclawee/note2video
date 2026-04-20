import type { WaveformHandle } from "./components/WaveformPanel";
import { WaveformPanel } from "./components/WaveformPanel";
import type { TimelineRow } from "@vitecut/timeline";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { FabricStage } from "./components/FabricStage";
import { TimelineSection } from "./components/TimelineSection";
import { canvasSize } from "./lib/canvasLayout";
import { activeCue, cuesForPage, loadProject } from "./lib/projectLoader";
import type { LoadedProject } from "./lib/projectTypes";

function buildRows(project: LoadedProject): TimelineRow[] {
  const imageRow: TimelineRow = {
    id: "slides",
    actions: project.slideClips.map((c) => ({
      id: `slide-${c.page}`,
      effectId: "image",
      start: c.start,
      end: Math.max(c.start + 0.2, c.end),
      flexible: true,
      movable: true,
      page: c.page,
      title: c.title,
    })),
  };
  const audioRow: TimelineRow = {
    id: "voice",
    actions: project.slideClips
      .filter((c) => c.audioAbs)
      .map((c) => ({
        id: `audio-${c.page}`,
        effectId: "audio",
        start: c.start,
        end: Math.max(c.start + 0.2, c.end),
        flexible: true,
        movable: true,
        page: c.page,
      })),
  };
  return [imageRow, audioRow];
}

export function App() {
  const waveRef = useRef<WaveformHandle>(null);

  const [project, setProject] = useState<LoadedProject | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const [editorData, setEditorData] = useState<TimelineRow[]>([]);
  const [duration, setDuration] = useState(120);
  const [currentTime, setCurrentTime] = useState(0);
  const [playing, setPlaying] = useState(false);
  const [zoom, setZoom] = useState(1);

  const [activePage, setActivePage] = useState(1);
  const [imageUrl, setImageUrl] = useState<string | null>(null);

  const logical = useMemo(() => {
    if (!project) {
      return { w: 1920, h: 1080 };
    }
    return canvasSize(project.manifest.ratio, project.manifest.resolution);
  }, [project]);

  const openDir = useCallback(async () => {
    setError(null);
    if (!window.n2v) {
      setError("请在 Electron 中运行（npm run dev），浏览器模式无法读取本地项目。");
      return;
    }
    setBusy(true);
    try {
      const dir = await window.n2v.openProjectDir();
      if (!dir) {
        return;
      }
      const p = await loadProject(dir);
      setProject(p);
      setEditorData(buildRows(p));
      setDuration(p.totalDuration);
      setCurrentTime(0);
      setPlaying(false);
      setActivePage(p.slideClips[0]?.page ?? 1);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      if (!project || !window.n2v) {
        setImageUrl(null);
        return;
      }
      const clip = project.slideClips.find((c) => c.page === activePage);
      if (!clip?.imageAbs) {
        setImageUrl(null);
        return;
      }
      try {
        const url = await window.n2v.toFileUrl(clip.imageAbs);
        if (!cancelled) {
          setImageUrl(url);
        }
      } catch {
        if (!cancelled) {
          setImageUrl(null);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [project, activePage]);

  useEffect(() => {
    if (!project) {
      return;
    }
    const t = currentTime;
    const clip =
      project.slideClips.find((c) => t >= c.start && t < c.end) ||
      project.slideClips.find((c) => t >= c.start) ||
      project.slideClips[0];
    if (clip && clip.page !== activePage) {
      setActivePage(clip.page);
    }
  }, [activePage, currentTime, project]);

  const cues = useMemo(() => (project ? cuesForPage(project, activePage) : []), [project, activePage]);
  const cueHit = useMemo(() => activeCue(cues, currentTime), [cues, currentTime]);
  const overlaySubtitle = cueHit?.text || (cues[0]?.text ?? "（无字幕 / 脚本）");

  const [audioUrl, setAudioUrl] = useState<string | null>(null);

  useEffect(() => {
    void (async () => {
      if (!project?.mergedAudioPath || !window.n2v) {
        setAudioUrl(null);
        return;
      }
      try {
        const u = await window.n2v.toFileUrl(project.mergedAudioPath);
        setAudioUrl(u);
      } catch {
        setAudioUrl(null);
      }
    })();
  }, [project]);

  const onWaveReady = useCallback((d: number) => {
    if (d > 0) {
      setDuration((prev) => Math.max(prev, d + 1));
    }
  }, []);

  const throttledTime = useCallback((t: number) => {
    setCurrentTime(t);
    if (project && t >= project.totalDuration - 0.05) {
      setPlaying(false);
    }
  }, [project]);

  useEffect(() => {
    if (!playing) {
      waveRef.current?.pause();
      return;
    }
    void waveRef.current?.play();
  }, [playing]);

  const activeTitle = useMemo(() => {
    if (!project) {
      return "Note2Video Studio";
    }
    const s = project.manifest.slides?.find((x) => Number(x.page) === activePage);
    return String(s?.title || `第 ${activePage} 页`);
  }, [activePage, project]);

  return (
    <div className="app-shell">
      <header className="app-toolbar">
        <h1>
          Note2Video Studio
          <span className="sub">Electron · Fabric · ViteCut · WaveSurfer</span>
        </h1>
        <button type="button" className="btn btn-primary" onClick={() => void openDir()} disabled={busy}>
          {busy ? "加载中…" : "打开项目目录…"}
        </button>
        {project && (
          <button
            type="button"
            className="btn"
            onClick={() => {
              void window.n2v?.openPath(project.rootDir);
            }}
          >
            在系统中打开
          </button>
        )}
        {project && audioUrl && (
          <>
            <button
              type="button"
              className="btn"
              onClick={() => {
                setPlaying((p) => !p);
              }}
            >
              {playing ? "暂停" : "播放"}
            </button>
            <button
              type="button"
              className="btn btn-ghost"
              onClick={() => {
                setPlaying(false);
                setCurrentTime(0);
              }}
            >
              回到起点
            </button>
          </>
        )}
        <div className="toolbar-spacer" />
        {error && <span className="status-warn">{error}</span>}
        {project && <span className="status-ok">已加载：{project.manifest.project_name || project.rootDir}</span>}
      </header>

      <div className="app-body">
        <div className="center-stage">
          {!project ? (
            <div className="empty-hint">
              选择由 <code>note2video build</code> 或各子命令生成的输出目录（含 <code>manifest.json</code>）。
              <br />
              本界面为<strong>独立工作室预览</strong>：时间轴与画布编辑为前端态；渲染成片仍以现有 CLI / Qt 流水线为准。
            </div>
          ) : (
            <>
              <div className="preview-wrap">
                <FabricStage
                  logicalW={logical.w}
                  logicalH={logical.h}
                  imageUrl={imageUrl}
                  overlayTitle={activeTitle}
                  overlaySubtitle={overlaySubtitle}
                />
              </div>
              <WaveformPanel
                ref={waveRef}
                audioUrl={audioUrl}
                currentTime={currentTime}
                onTimeUpdate={throttledTime}
                onReady={onWaveReady}
              />
              <TimelineSection
                duration={duration}
                editorData={editorData}
                onEditorDataChange={setEditorData}
                currentTime={currentTime}
                playing={playing}
                onCurrentTimeChange={(t) => {
                  setCurrentTime(t);
                }}
                zoom={zoom}
                onZoomChange={setZoom}
              />
            </>
          )}
        </div>

        {project && (
          <aside className="inspector">
            <h2>项目</h2>
            <dl>
              <dt>比例 / 分辨率</dt>
              <dd>
                {project.manifest.ratio || "16:9"} · {project.manifest.resolution || "1080p"}
              </dd>
              <dt>画布逻辑像素</dt>
              <dd>
                {logical.w} × {logical.h}
              </dd>
              <dt>当前时间</dt>
              <dd>{currentTime.toFixed(3)} s</dd>
              <dt>当前页</dt>
              <dd>{activePage}</dd>
            </dl>
            <h2>页面</h2>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {project.slideClips.map((c) => (
                <button
                  key={c.page}
                  type="button"
                  className="btn"
                  style={{
                    textAlign: "left",
                    borderColor: c.page === activePage ? "rgba(108,140,255,0.55)" : undefined,
                    background: c.page === activePage ? "rgba(108,140,255,0.12)" : undefined,
                  }}
                  onClick={() => {
                    setActivePage(c.page);
                    setCurrentTime(c.start);
                  }}
                >
                  P{c.page} · {c.title}
                </button>
              ))}
            </div>
            <h2>提示</h2>
            <p style={{ fontSize: 12, color: "#8b93a7", lineHeight: 1.5, margin: 0 }}>
              拖拽时间轴上的幻灯片 / 音频块可调整相对位置（前端编辑）。导出到成片请仍使用原有 render 命令。
            </p>
          </aside>
        )}
      </div>
    </div>
  );
}
