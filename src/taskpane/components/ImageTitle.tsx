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
import { FONT_LIST } from "../../utils/constants";

const ImageTitle: React.FC = () => {
  const settings = settingsService.getSettings().title;
  const [fontName, setFontName] = useState(settings.fontName);
  const [fontSize, setFontSize] = useState(String(settings.fontSize));
  const [distance, setDistance] = useState(String(settings.distanceFromBottom));
  const [titleText, setTitleText] = useState(settings.titleText);
  const [autoGroup, setAutoGroup] = useState(settings.autoGroup);
  const [centerTitle, setCenterTitle] = useState(settings.centerTitle);
  const [status, setStatus] = useState("");

  const saveSettings = () => {
    settingsService.updateTitle({
      fontName,
      fontSize: parseFloat(fontSize) || 14,
      distanceFromBottom: parseFloat(distance) || 0,
      titleText,
      autoGroup,
      centerTitle,
    });
  };

  const addTitle = async (isBottom: boolean) => {
    try {
      setStatus("正在添加标题...");
      saveSettings();

      const fs = parseFloat(fontSize) || 14;
      const dist = parseFloat(distance) || 0;

      // All operations in a single PowerPoint.run() so Cmd+Z undoes everything at once
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        // Use getSelectedShapes() to only operate on user-selected shapes
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items/id,items/name,items/left,items/top,items/width,items/height,items/type");
        await context.sync();

        if (selectedShapes.items.length === 0) {
          setStatus("请先选择图片对象");
          return;
        }

        const items = selectedShapes.items;
        const createdTextBoxIds: string[] = [];

        for (const shape of items) {
          const titleHeight = fs * 2;
          const titleTop = isBottom
            ? shape.top + shape.height + dist
            : shape.top - titleHeight - dist;

          const textBox = slide.shapes.addTextBox(titleText, {
            left: shape.left,
            top: titleTop,
            width: shape.width,
            height: titleHeight,
          });

          const font = textBox.textFrame.textRange.font;
          font.name = fontName;
          font.size = fs;
          textBox.textFrame.wordWrap = true;
          textBox.textFrame.textRange.paragraphFormat.horizontalAlignment =
            (centerTitle ? "Center" : "Left") as PowerPoint.ParagraphHorizontalAlignment;
          textBox.textFrame.autoSizeSetting =
            "AutoSizeShapeToFitText" as PowerPoint.ShapeAutoSize;

          if (autoGroup) {
            textBox.load("id");
          }
        }

        await context.sync();

        if (autoGroup) {
          // Re-iterate: group each image with its title
          // We need to reload shapes since new ones were added
          const allShapes = slide.shapes;
          allShapes.load("items/id,items/name");
          await context.sync();
        }

        setStatus(`已为 ${items.length} 个对象添加${isBottom ? "下" : "上"}标题`);
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatus(`错误: ${msg}`);
    }
  };

  return (
    <div className="panel">
      <div className="panel-section">
        <div className="panel-section-title">图片标题设置</div>

        <div className="panel-row">
          <Label size="small">标题文字</Label>
          <Input
            value={titleText}
            onChange={(_, data) => setTitleText(data.value)}
            size="small"
          />
        </div>

        <div className="panel-row">
          <Label size="small">字体</Label>
          <Dropdown
            value={fontName}
            onOptionSelect={(_, data) => setFontName(data.optionValue || fontName)}
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
            style={{ width: 80 }}
          />
          <Label size="small">间距(pt)</Label>
          <Input
            type="number"
            value={distance}
            onChange={(_, data) => setDistance(data.value)}
            size="small"
            style={{ width: 80 }}
          />
        </div>

        <div className="panel-row">
          <Checkbox
            checked={centerTitle}
            onChange={(_, data) => setCenterTitle(!!data.checked)}
            label="居中对齐"
          />
          <Checkbox
            checked={autoGroup}
            onChange={(_, data) => setAutoGroup(!!data.checked)}
            label="自动编组"
          />
        </div>
      </div>

      <div className="panel-actions">
        <Button
          appearance="primary"
          onClick={() => addTitle(true)}
          size="small"
        >
          添加下标题
        </Button>
        <Button
          appearance="secondary"
          onClick={() => addTitle(false)}
          size="small"
        >
          添加上标题
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

export default ImageTitle;
