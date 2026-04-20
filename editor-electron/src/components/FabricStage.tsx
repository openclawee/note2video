import { Canvas, FabricImage, FabricText, Rect, Shadow } from "fabric";
import { useCallback, useEffect, useRef, useState } from "react";

type Props = {
  logicalW: number;
  logicalH: number;
  imageUrl: string | null;
  overlayTitle: string;
  overlaySubtitle: string;
};

const PAD = 48;

export function FabricStage({ logicalW, logicalH, imageUrl, overlayTitle, overlaySubtitle }: Props) {
  const hostRef = useRef<HTMLDivElement>(null);
  const canvasElRef = useRef<HTMLCanvasElement>(null);
  const fabricRef = useRef<Canvas | null>(null);
  const [box, setBox] = useState({ w: 640, h: 360 });

  const measure = useCallback(() => {
    const el = hostRef.current;
    if (!el) {
      return;
    }
    const maxW = Math.max(200, el.clientWidth - PAD);
    const maxH = Math.max(200, el.clientHeight - PAD);
    const scale = Math.min(maxW / logicalW, maxH / logicalH, 1);
    setBox({
      w: Math.floor(logicalW * scale),
      h: Math.floor(logicalH * scale),
    });
  }, [logicalH, logicalW]);

  useEffect(() => {
    measure();
    const ro = new ResizeObserver(() => measure());
    if (hostRef.current) {
      ro.observe(hostRef.current);
    }
    return () => ro.disconnect();
  }, [measure]);

  useEffect(() => {
    const el = canvasElRef.current;
    if (!el) {
      return;
    }
    const canvas = new Canvas(el, {
      width: box.w,
      height: box.h,
      backgroundColor: "#050608",
      selection: true,
      preserveObjectStacking: true,
    });
    fabricRef.current = canvas;

    const bottomBar = new Rect({
      left: 0,
      top: box.h - 56,
      width: box.w,
      height: 56,
      fill: "rgba(0,0,0,0.55)",
      selectable: false,
      evented: false,
    });
    canvas.add(bottomBar);

    return () => {
      canvas.dispose();
      fabricRef.current = null;
    };
  }, [box.w, box.h]);

  useEffect(() => {
    const canvas = fabricRef.current;
    if (!canvas) {
      return;
    }

    const scale = box.w / logicalW;

    void (async () => {
      canvas.getObjects().forEach((o) => {
        if ((o as { n2vRole?: string }).n2vRole) {
          canvas.remove(o);
        }
      });

      if (imageUrl) {
        try {
          const img = await FabricImage.fromURL(imageUrl, { crossOrigin: "anonymous" });
          img.set({
            originX: "center",
            originY: "center",
            left: box.w / 2,
            top: box.h / 2 - 8,
            selectable: true,
            hasControls: true,
            lockRotation: false,
            cornerStyle: "circle",
            borderColor: "#6c8cff",
            cornerColor: "#6c8cff",
          });
          const fit = Math.min((box.w - 24) / (img.width || 1), (box.h - 80) / (img.height || 1));
          img.scale(fit);
          (img as unknown as { n2vRole?: string }).n2vRole = "slide";
          canvas.add(img);
          canvas.sendObjectToBack(img);
        } catch {
          const placeholder = new FabricText("无法加载幻灯片图片", {
            left: box.w / 2,
            top: box.h / 2,
            originX: "center",
            originY: "center",
            fill: "#7a8499",
            fontSize: 14 * scale,
            fontFamily: "Inter, system-ui, sans-serif",
          });
          (placeholder as unknown as { n2vRole?: string }).n2vRole = "placeholder";
          canvas.add(placeholder);
        }
      } else {
        const placeholder = new FabricText("未找到当前页图片\n请先在 CLI / Qt 中运行 extract", {
          left: box.w / 2,
          top: box.h / 2,
          originX: "center",
          originY: "center",
          fill: "#7a8499",
          fontSize: 13 * scale,
          fontFamily: "Inter, system-ui, sans-serif",
          textAlign: "center",
        });
        (placeholder as unknown as { n2vRole?: string }).n2vRole = "placeholder";
        canvas.add(placeholder);
      }

      const title = new FabricText(overlayTitle || " ", {
        left: 16,
        top: 14 * scale,
        fill: "#ffffff",
        fontSize: Math.max(13, 16 * scale),
        fontWeight: "600",
        fontFamily: "Inter, system-ui, sans-serif",
        shadow: new Shadow({ color: "rgba(0,0,0,0.65)", blur: 6, offsetX: 0, offsetY: 2 }),
        selectable: true,
        hasControls: true,
      });
      (title as unknown as { n2vRole?: string }).n2vRole = "title";

      const badge = new FabricText("WYSIWYG", {
        left: box.w - 16,
        top: 14 * scale,
        originX: "right",
        fill: "#c8d6ff",
        fontSize: 11 * scale,
        fontFamily: "Inter, system-ui, sans-serif",
        backgroundColor: "rgba(108,140,255,0.2)",
        padding: 6,
        selectable: false,
      });
      (badge as unknown as { n2vRole?: string }).n2vRole = "badge";

      const sub = new FabricText(overlaySubtitle || " ", {
        left: box.w / 2,
        top: box.h - 28,
        originX: "center",
        originY: "center",
        fill: "#f2f4ff",
        fontSize: Math.max(12, 15 * scale),
        fontFamily: "Inter, system-ui, sans-serif",
        textAlign: "center",
        lineHeight: 1.25,
        shadow: new Shadow({ color: "rgba(0,0,0,0.85)", blur: 4, offsetX: 0, offsetY: 2 }),
        width: box.w - 48,
        splitByGrapheme: true,
        selectable: true,
        hasControls: true,
      });
      (sub as unknown as { n2vRole?: string }).n2vRole = "subtitle";

      canvas.add(title, badge, sub);

      const bar = canvas.getObjects().find((o) => o.type === "rect" && !(o as unknown as { n2vRole?: string }).n2vRole);
      if (bar) {
        canvas.sendObjectToBack(bar);
      }
      canvas.requestRenderAll();
    })();
  }, [box.h, box.w, imageUrl, logicalH, logicalW, overlaySubtitle, overlayTitle]);

  return (
    <div ref={hostRef} className="preview-canvas-host" style={{ width: "100%", height: "100%" }}>
      <canvas ref={canvasElRef} />
    </div>
  );
}
