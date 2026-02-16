import React, { useState } from "react";
import { Button } from "@fluentui/react-components";

type ShapeType = "Image" | "TextBox";

const SelectShapes: React.FC = () => {
  const [status, setStatus] = useState("");

  const selectAllByType = async (shapeType: ShapeType) => {
    const label = shapeType === "Image" ? "图片" : "文本框";
    try {
      setStatus(`正在选中${label}...`);

      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const shapes = slide.shapes;
        shapes.load("items/id,items/type");
        await context.sync();

        const ids = shapes.items
          .filter((s) => String(s.type) === shapeType)
          .map((s) => s.id);

        if (ids.length === 0) {
          setStatus(`当前页面没有${label}对象`);
          return;
        }

        slide.setSelectedShapes(ids);
        await context.sync();
        setStatus(`已选中 ${ids.length} 个${label}`);
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatus(`错误: ${msg}`);
    }
  };

  const filterSelectedByType = async (shapeType: ShapeType) => {
    const label = shapeType === "Image" ? "图片" : "文本框";
    try {
      setStatus(`正在筛选${label}...`);

      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        const selected = context.presentation.getSelectedShapes();
        selected.load("items/id,items/type");
        await context.sync();

        if (selected.items.length === 0) {
          setStatus("请先选择对象");
          return;
        }

        const ids = selected.items
          .filter((s) => String(s.type) === shapeType)
          .map((s) => s.id);

        if (ids.length === 0) {
          setStatus(`选中的对象中没有${label}`);
          return;
        }

        slide.setSelectedShapes(ids);
        await context.sync();
        setStatus(`已筛选出 ${ids.length} 个${label}`);
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatus(`错误: ${msg}`);
    }
  };

  return (
    <div className="panel">
      <div className="panel-section">
        <div className="panel-section-title">全页选择</div>
        <div className="panel-actions">
          <Button
            appearance="primary"
            onClick={() => selectAllByType("Image")}
            size="small"
          >
            选中所有图片
          </Button>
          <Button
            appearance="primary"
            onClick={() => selectAllByType("TextBox")}
            size="small"
          >
            选中所有文本框
          </Button>
        </div>
      </div>

      <div className="panel-section">
        <div className="panel-section-title">筛选选择</div>
        <div className="panel-actions">
          <Button
            appearance="secondary"
            onClick={() => filterSelectedByType("Image")}
            size="small"
          >
            筛选图片
          </Button>
          <Button
            appearance="secondary"
            onClick={() => filterSelectedByType("TextBox")}
            size="small"
          >
            筛选文本框
          </Button>
        </div>
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

export default SelectShapes;
