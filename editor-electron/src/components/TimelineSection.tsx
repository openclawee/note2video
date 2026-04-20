import { Timeline, type TimelineRow } from "@vitecut/timeline";
import "@vitecut/timeline/style.css";
import { useMemo, useRef } from "react";

type Props = {
  duration: number;
  editorData: TimelineRow[];
  onEditorDataChange: (rows: TimelineRow[]) => void;
  currentTime: number;
  playing: boolean;
  onCurrentTimeChange: (t: number) => void;
  zoom: number;
  onZoomChange: (z: number) => void;
};

export function TimelineSection({
  duration,
  editorData,
  onEditorDataChange,
  currentTime,
  playing,
  onCurrentTimeChange,
  zoom,
  onZoomChange,
}: Props) {
  const timelineRef = useRef<import("@vitecut/timeline").TimelineState | null>(null);

  const safeDuration = useMemo(() => Math.max(10, duration), [duration]);

  return (
    <div className="timeline-panel">
      <div className="timeline-toolbar">
        <span className="badge">@vitecut/timeline</span>
        <span style={{ fontSize: 12, color: "#8b93a7" }}>缩放</span>
        <input
          type="range"
          min={0.35}
          max={2.5}
          step={0.05}
          value={zoom}
          onChange={(e) => onZoomChange(Number(e.target.value))}
        />
        <button type="button" className="btn btn-ghost" onClick={() => onZoomChange(1)}>
          重置缩放
        </button>
        <span style={{ fontSize: 12, color: "#8b93a7", marginLeft: "auto" }}>
          t = {currentTime.toFixed(2)}s / {safeDuration.toFixed(1)}s
        </span>
      </div>
      <div className="timeline-host vitecut-timeline-host">
        <Timeline
          ref={timelineRef}
          editorData={editorData}
          duration={safeDuration}
          playing={playing}
          currentTime={currentTime}
          zoom={zoom}
          onZoomChange={onZoomChange}
          onEditorDataChange={onEditorDataChange}
          onCursorDrag={onCurrentTimeChange}
          onCursorDragEnd={onCurrentTimeChange}
          onClickTimeArea={(nextTime) => {
            onCurrentTimeChange(nextTime);
            return true;
          }}
          dragSnapToClipEdges
          trimSnapToClipEdges
          trimSnapToTimelineTicks
          trackHeightPresets={{ video: 52, audio: 44, main: 52 }}
          getActionRender={(action) => {
            const label = String(action.page ?? action.effectId ?? "");
            return (
              <span style={{ fontSize: 11, padding: "4px 8px", color: "#e8eaf2", fontWeight: 500 }}>
                {action.effectId === "image" ? `P${label}` : action.effectId === "audio" ? "VO" : String(action.effectId)}
              </span>
            );
          }}
        />
      </div>
    </div>
  );
}
