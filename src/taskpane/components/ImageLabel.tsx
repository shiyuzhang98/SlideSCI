import React, { useState } from "react";
import {
  Button,
  Input,
  Label,
  Checkbox,
  Dropdown,
  Option,
} from "@fluentui/react-components";
import { settingsService } from "../../services/settings-service";
import { shapeService } from "../../services/shape-service";
import { generateLabel } from "../../core/label-generator";
import { groupAndSortShapes } from "../../core/shape-sorting";
import { FONT_LIST, LABEL_TEMPLATES } from "../../utils/constants";

const ImageLabel: React.FC = () => {
  const settings = settingsService.getSettings().label;
  const [offsetX, setOffsetX] = useState(String(settings.offsetX));
  const [offsetY, setOffsetY] = useState(String(settings.offsetY));
  const [template, setTemplate] = useState(settings.template);
  const [fontName, setFontName] = useState(settings.fontName);
  const [fontSize, setFontSize] = useState(String(settings.fontSize));
  const [bold, setBold] = useState(settings.bold);
  const [startIndex, setStartIndex] = useState(String(settings.startIndex));
  const [autoUpdate, setAutoUpdate] = useState(settings.autoUpdateIndex);
  const [status, setStatus] = useState("");

  const saveSettings = () => {
    settingsService.updateLabel({
      offsetX: parseFloat(offsetX) || -20,
      offsetY: parseFloat(offsetY) || -7,
      template,
      fontName,
      fontSize: parseFloat(fontSize) || 12,
      bold,
      startIndex: parseInt(startIndex) || 1,
      autoUpdateIndex: autoUpdate,
    });
  };

  const addLabels = async () => {
    try {
      setStatus("正在添加标签...");
      saveSettings();

      const fs = parseFloat(fontSize) || 12;
      const ox = parseFloat(offsetX) || -20;
      const oy = parseFloat(offsetY) || -7;
      const si = parseInt(startIndex) || 1;

      // All operations in a single PowerPoint.run() so Cmd+Z undoes everything at once
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        // Use getSelectedShapes() to only operate on user-selected shapes
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items/id,items/name,items/left,items/top,items/width,items/height,items/type");
        await context.sync();

        const items = selectedShapes.items.map((s) => ({
          id: s.id,
          name: s.name,
          left: s.left,
          top: s.top,
          width: s.width,
          height: s.height,
          type: s.type as string,
        }));

        if (items.length === 0) {
          setStatus("请先选择图片对象");
          return;
        }

        // Sort shapes using position-based grouping (same as original C#)
        const sortedShapes = groupAndSortShapes(items);

        for (let i = 0; i < sortedShapes.length; i++) {
          const shape = sortedShapes[i];
          const label = generateLabel(template, i, si);

          const textBox = slide.shapes.addTextBox(label, {
            left: shape.left + ox,
            top: shape.top + oy,
            width: 0,
            height: fs * 2,
          });

          const font = textBox.textFrame.textRange.font;
          font.name = fontName;
          font.size = fs;
          if (bold) font.bold = true;
          textBox.textFrame.textRange.paragraphFormat.horizontalAlignment =
            "Left" as PowerPoint.ParagraphHorizontalAlignment;
          textBox.textFrame.autoSizeSetting =
            "AutoSizeShapeToFitText" as PowerPoint.ShapeAutoSize;
          textBox.textFrame.wordWrap = false;
        }

        await context.sync();

        // Auto-update start index for next batch
        if (autoUpdate) {
          const nextIndex = si + sortedShapes.length;
          setStartIndex(String(nextIndex));
          settingsService.updateLabel({ startIndex: nextIndex });
        }

        setStatus(`已为 ${sortedShapes.length} 个对象添加标签`);
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatus(`错误: ${msg}`);
    }
  };

  return (
    <div className="panel">
      <div className="panel-section">
        <div className="panel-section-title">图片标签设置</div>

        <div className="panel-row">
          <Label size="small">标签格式</Label>
          <Dropdown
            value={
              LABEL_TEMPLATES.find((t) => t.key === template)?.label ||
              "A, B, C"
            }
            onOptionSelect={(_, data) =>
              setTemplate(data.optionValue || "A")
            }
            size="small"
          >
            {LABEL_TEMPLATES.map((t) => (
              <Option key={t.key} value={t.key}>
                {t.label}
              </Option>
            ))}
          </Dropdown>
        </div>

        <div className="panel-row">
          <Label size="small">字体</Label>
          <Dropdown
            value={fontName}
            onOptionSelect={(_, data) =>
              setFontName(data.optionValue || fontName)
            }
            size="small"
          >
            {FONT_LIST.map((f) => (
              <Option key={f} value={f}>
                {f}
              </Option>
            ))}
          </Dropdown>
        </div>

        <div className="panel-row">
          <Label size="small">字号</Label>
          <Input
            type="number"
            value={fontSize}
            onChange={(_, data) => setFontSize(data.value)}
            size="small"
            style={{ width: 70 }}
          />
          <Label size="small">起始</Label>
          <Input
            type="number"
            value={startIndex}
            onChange={(_, data) => setStartIndex(data.value)}
            size="small"
            style={{ width: 60 }}
          />
        </div>

        <div className="panel-row">
          <Label size="small">X偏移</Label>
          <Input
            type="number"
            value={offsetX}
            onChange={(_, data) => setOffsetX(data.value)}
            size="small"
            style={{ width: 70 }}
          />
          <Label size="small">Y偏移</Label>
          <Input
            type="number"
            value={offsetY}
            onChange={(_, data) => setOffsetY(data.value)}
            size="small"
            style={{ width: 70 }}
          />
        </div>

        <div className="panel-row">
          <Checkbox
            checked={bold}
            onChange={(_, data) => setBold(!!data.checked)}
            label="粗体"
          />
          <Checkbox
            checked={autoUpdate}
            onChange={(_, data) => setAutoUpdate(!!data.checked)}
            label="自动更新编号"
          />
        </div>
      </div>

      <div className="panel-actions">
        <Button appearance="primary" onClick={addLabels} size="small">
          添加标签
        </Button>
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

export default ImageLabel;
