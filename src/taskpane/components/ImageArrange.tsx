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
import { groupAndSortShapes } from "../../core/shape-sorting";
import { arrangeShapes } from "../../core/grid-layout";
import { ARRANGE_MODES, SORT_MODES } from "../../utils/constants";

const ImageArrange: React.FC = () => {
  const settings = settingsService.getSettings().arrange;
  const [sortType, setSortType] = useState(settings.sortType);
  const [alignType, setAlignType] = useState(settings.alignType);
  const [colNum, setColNum] = useState(String(settings.colNum));
  const [colSpace, setColSpace] = useState(String(settings.colSpace));
  const [rowSpace, setRowSpace] = useState(String(settings.rowSpace));
  const [imgWidth, setImgWidth] = useState(settings.imgWidth);
  const [imgHeight, setImgHeight] = useState(settings.imgHeight);
  const [excludeText, setExcludeText] = useState(settings.excludeTextShapes);
  const [status, setStatus] = useState("");

  const saveSettings = () => {
    settingsService.updateArrange({
      sortType,
      alignType,
      colNum: parseInt(colNum) || 3,
      colSpace: parseFloat(colSpace) || 10,
      rowSpace: parseFloat(rowSpace) || 25,
      imgWidth,
      imgHeight,
      excludeTextShapes: excludeText,
    });
  };

  const doArrange = async () => {
    try {
      setStatus("正在排列图片...");
      saveSettings();

      const cn = parseInt(colNum) || 3;
      const cs = parseFloat(colSpace) || 10;
      const rs = parseFloat(rowSpace) || 25;
      const cw = parseFloat(imgWidth) || 0; // cm
      const ch = parseFloat(imgHeight) || 0; // cm

      const shapes = await shapeService.getSelectedShapes(excludeText);

      if (shapes.length === 0) {
        setStatus("请先选择图片对象");
        return;
      }

      // Sort shapes
      const sortedShapes =
        sortType === 0 ? groupAndSortShapes(shapes) : shapes;

      // Calculate arrangement
      const updates = arrangeShapes(sortedShapes, alignType, {
        colNum: cn,
        colSpace: cs,
        rowSpace: rs,
        customWidth: cw,
        customHeight: ch,
      });

      // Apply positions via PowerPoint API
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        for (const update of updates) {
          const shape = slide.shapes.getItem(update.id);
          shape.left = update.left;
          shape.top = update.top;
          shape.width = update.width;
          shape.height = update.height;
        }
        await context.sync();
      });

      setStatus(`已排列 ${updates.length} 个对象`);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatus(`错误: ${msg}`);
    }
  };

  return (
    <div className="panel">
      <div className="panel-section">
        <div className="panel-section-title">图片排列设置</div>

        <div className="panel-row">
          <Label size="small">排序方式</Label>
          <Dropdown
            value={SORT_MODES.find((m) => m.key === sortType)?.label || "按位置排序"}
            onOptionSelect={(_, data) =>
              setSortType(parseInt(data.optionValue || "0"))
            }
            size="small"
          >
            {SORT_MODES.map((m) => (
              <Option key={m.key} value={String(m.key)}>
                {m.label}
              </Option>
            ))}
          </Dropdown>
        </div>

        <div className="panel-row">
          <Label size="small">排列模式</Label>
          <Dropdown
            value={
              ARRANGE_MODES.find((m) => m.key === alignType)?.label ||
              "列最大宽度排列"
            }
            onOptionSelect={(_, data) =>
              setAlignType(parseInt(data.optionValue || "0"))
            }
            size="small"
          >
            {ARRANGE_MODES.map((m) => (
              <Option key={m.key} value={String(m.key)}>
                {m.label}
              </Option>
            ))}
          </Dropdown>
        </div>

        <div className="panel-row">
          <Label size="small">列数</Label>
          <Input
            type="number"
            value={colNum}
            onChange={(_, data) => setColNum(data.value)}
            size="small"
            style={{ width: 60 }}
          />
          <Label size="small">列间距</Label>
          <Input
            type="number"
            value={colSpace}
            onChange={(_, data) => setColSpace(data.value)}
            size="small"
            style={{ width: 60 }}
          />
        </div>

        <div className="panel-row">
          <Label size="small">行间距</Label>
          <Input
            type="number"
            value={rowSpace}
            onChange={(_, data) => setRowSpace(data.value)}
            size="small"
            style={{ width: 60 }}
          />
        </div>

        <div className="panel-row">
          <Label size="small">宽(cm)</Label>
          <Input
            value={imgWidth}
            onChange={(_, data) => setImgWidth(data.value)}
            size="small"
            placeholder="自动"
            style={{ width: 70 }}
          />
          <Label size="small">高(cm)</Label>
          <Input
            value={imgHeight}
            onChange={(_, data) => setImgHeight(data.value)}
            size="small"
            placeholder="自动"
            style={{ width: 70 }}
          />
        </div>

        <div className="panel-row">
          <Checkbox
            checked={excludeText}
            onChange={(_, data) => setExcludeText(!!data.checked)}
            label="排除文本框"
          />
        </div>
      </div>

      <div className="panel-actions">
        <Button appearance="primary" onClick={doArrange} size="small">
          排列图片
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

export default ImageArrange;
