import React, { useState, useRef } from "react";
import { Button, Divider } from "@fluentui/react-components";
import {
  CopyRegular,
  ClipboardPasteRegular,
} from "@fluentui/react-icons";
import { shapeService } from "../../services/shape-service";
import { CopiedPosition, CopiedDimensions, FontProperties } from "../../models/types";

const FormatCopy: React.FC = () => {
  const [status, setStatus] = useState("");
  const copiedPositions = useRef<CopiedPosition[]>([]);
  const copiedDimensions = useRef<CopiedDimensions | null>(null);
  const copiedWidth = useRef<number>(0);
  const copiedHeight = useRef<number>(0);
  const copiedFont = useRef<FontProperties | null>(null);

  // ========== Position ==========
  const copyPosition = async () => {
    try {
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }
      copiedPositions.current = shapes.map((s) => ({
        centerX: s.left + s.width / 2,
        centerY: s.top + s.height / 2,
      }));
      setStatus(`已复制 ${shapes.length} 个对象的位置`);
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  const pastePosition = async () => {
    try {
      if (copiedPositions.current.length === 0) {
        setStatus("请先复制位置");
        return;
      }
      const shapes = await shapeService.getSelectedShapes();
      const count = Math.min(shapes.length, copiedPositions.current.length);
      if (count === 0) {
        setStatus("请先选择对象");
        return;
      }

      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapeRefs = [];
        for (let i = 0; i < count; i++) {
          const shape = slide.shapes.getItem(shapes[i].id);
          shape.load("width,height");
          shapeRefs.push(shape);
        }
        await context.sync();
        for (let i = 0; i < count; i++) {
          shapeRefs[i].left = copiedPositions.current[i].centerX - shapeRefs[i].width / 2;
          shapeRefs[i].top = copiedPositions.current[i].centerY - shapeRefs[i].height / 2;
        }
        await context.sync();
      });

      setStatus(`已粘贴位置到 ${count} 个对象`);
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  // ========== Width ==========
  const copyWidth = async () => {
    try {
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }
      copiedWidth.current = shapes[0].width;
      setStatus("已复制宽度");
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  const pasteWidth = async () => {
    try {
      if (copiedWidth.current <= 0) {
        setStatus("请先复制宽度");
        return;
      }
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }

      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapeRefs = shapes.map((s) => {
          const shape = slide.shapes.getItem(s.id);
          shape.load("width,height");
          return shape;
        });
        await context.sync();
        for (const shape of shapeRefs) {
          const ratio = shape.height / shape.width;
          shape.width = copiedWidth.current;
          shape.height = copiedWidth.current * ratio;
        }
        await context.sync();
      });

      setStatus(`已粘贴宽度到 ${shapes.length} 个对象`);
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  // ========== Height ==========
  const copyHeight = async () => {
    try {
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }
      copiedHeight.current = shapes[0].height;
      setStatus("已复制高度");
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  const pasteHeight = async () => {
    try {
      if (copiedHeight.current <= 0) {
        setStatus("请先复制高度");
        return;
      }
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }

      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapeRefs = shapes.map((s) => {
          const shape = slide.shapes.getItem(s.id);
          shape.load("width,height");
          return shape;
        });
        await context.sync();
        for (const shape of shapeRefs) {
          const ratio = shape.width / shape.height;
          shape.height = copiedHeight.current;
          shape.width = copiedHeight.current * ratio;
        }
        await context.sync();
      });

      setStatus(`已粘贴高度到 ${shapes.length} 个对象`);
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  // ========== Width + Height ==========
  const copyDimensions = async () => {
    try {
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }
      copiedDimensions.current = {
        width: shapes[0].width,
        height: shapes[0].height,
      };
      setStatus("已复制宽高");
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  const pasteDimensions = async () => {
    try {
      if (!copiedDimensions.current) {
        setStatus("请先复制宽高");
        return;
      }
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }

      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        for (const s of shapes) {
          const shape = slide.shapes.getItem(s.id);
          shape.width = copiedDimensions.current!.width;
          shape.height = copiedDimensions.current!.height;
        }
        await context.sync();
      });

      setStatus(`已粘贴宽高到 ${shapes.length} 个对象`);
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  // ========== Format (Font) ==========
  const copyFormat = async () => {
    try {
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }
      const font = await shapeService.getShapeFont(shapes[0].id);
      if (font) {
        copiedFont.current = font;
        setStatus("已复制格式");
      } else {
        setStatus("无法读取格式");
      }
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  const pasteFormat = async () => {
    try {
      if (!copiedFont.current) {
        setStatus("请先复制格式");
        return;
      }
      const shapes = await shapeService.getSelectedShapes();
      if (shapes.length === 0) {
        setStatus("请先选择对象");
        return;
      }

      for (const s of shapes) {
        await shapeService.setShapeFont(s.id, copiedFont.current);
      }

      setStatus(`已粘贴格式到 ${shapes.length} 个对象`);
    } catch (err: unknown) {
      setStatus(`错误: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  const renderCopyPasteRow = (
    label: string,
    onCopy: () => void,
    onPaste: () => void
  ) => (
    <div className="panel-row" style={{ justifyContent: "space-between" }}>
      <span style={{ fontSize: 12, minWidth: 50 }}>{label}</span>
      <div style={{ display: "flex", gap: 4 }}>
        <Button
          appearance="subtle"
          icon={<CopyRegular />}
          onClick={onCopy}
          size="small"
        >
          复制
        </Button>
        <Button
          appearance="subtle"
          icon={<ClipboardPasteRegular />}
          onClick={onPaste}
          size="small"
        >
          粘贴
        </Button>
      </div>
    </div>
  );

  return (
    <div className="panel">
      <div className="panel-section">
        <div className="panel-section-title">位置与尺寸</div>
        {renderCopyPasteRow("位置", copyPosition, pastePosition)}
        {renderCopyPasteRow("宽度", copyWidth, pasteWidth)}
        {renderCopyPasteRow("高度", copyHeight, pasteHeight)}
        {renderCopyPasteRow("宽高", copyDimensions, pasteDimensions)}
      </div>

      <Divider style={{ margin: "4px 0" }} />

      <div className="panel-section">
        <div className="panel-section-title">格式</div>
        {renderCopyPasteRow("文字格式", copyFormat, pasteFormat)}
      </div>

      {status && (
        <div
          className={`status-message ${
            status.startsWith("错误") ? "error" : "success"
          }`}
        >
          {status}
        </div>
      )}
    </div>
  );
};

export default FormatCopy;
