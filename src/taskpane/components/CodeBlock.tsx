import React, { useState } from "react";
import {
  Button,
  Dropdown,
  Option,
  Label,
  Textarea,
  Switch,
  Input,
} from "@fluentui/react-components";
import { settingsService } from "../../services/settings-service";
import {
  highlightCode,
  getThemeBackground,
} from "../../core/code-highlighter";
import { CODE_LANGUAGES } from "../../utils/constants";

const CodeBlock: React.FC = () => {
  const settings = settingsService.getSettings().codeBlock;
  const [language, setLanguage] = useState(settings.language);
  const [darkTheme, setDarkTheme] = useState(settings.darkTheme);
  const [code, setCode] = useState("");
  const [fontSize, setFontSize] = useState("12");
  const [status, setStatus] = useState("");

  const saveSettings = () => {
    settingsService.updateCodeBlock({ language, darkTheme });
  };

  const insertCodeBlock = async () => {
    if (!code.trim()) {
      setStatus("请输入代码");
      return;
    }

    try {
      setStatus("正在插入代码块...");
      saveSettings();

      const fs = parseFloat(fontSize) || 12;
      const theme = getThemeBackground(darkTheme);
      const tokens = highlightCode(code, language, darkTheme);

      await PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);

        // Create text box with the code
        const textBox = slide.shapes.addTextBox(code, {
          left: 100,
          top: 100,
          width: 500,
          height: 300,
        });

        // Set background fill
        textBox.fill.setSolidColor(theme.background);

        // Set border
        textBox.lineFormat.color = theme.border;
        textBox.lineFormat.weight = 1;

        // Set base font formatting
        const textRange = textBox.textFrame.textRange;
        const font = textRange.font;
        font.name = "Consolas";
        font.size = fs;
        font.color = theme.foreground;
        textRange.paragraphFormat.horizontalAlignment =
          "Left" as PowerPoint.ParagraphHorizontalAlignment;

        // Set margins (properties exist at runtime but not in TS types)
        const frame = textBox.textFrame as PowerPoint.TextFrame & {
          marginLeft: number;
          marginRight: number;
          marginTop: number;
          marginBottom: number;
        };
        frame.marginLeft = 10;
        frame.marginRight = 10;
        frame.marginTop = 5;
        frame.marginBottom = 5;

        // Apply per-token syntax highlighting
        for (const token of tokens) {
          try {
            const subRange = textRange.getSubstring(
              token.start,
              token.length
            );
            subRange.font.color = token.color;
          } catch {
            // Skip tokens that fail (edge cases)
          }
        }

        // Auto-size to fit content
        textBox.textFrame.autoSizeSetting =
          "AutoSizeShapeToFitText" as PowerPoint.ShapeAutoSize;
        textBox.textFrame.wordWrap = false;

        await context.sync();
        setStatus("代码块已插入");
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatus(`错误: ${msg}`);
    }
  };

  // Generate preview tokens for the preview pane
  const previewTokens = code ? highlightCode(code, language, darkTheme) : [];
  const theme = getThemeBackground(darkTheme);

  // Build preview: split code into colored segments
  const renderPreview = () => {
    if (!code) return null;

    // Create an array of characters with their colors
    const charColors = new Array(code.length).fill(theme.foreground);
    for (const token of previewTokens) {
      for (let i = token.start; i < token.start + token.length && i < code.length; i++) {
        charColors[i] = token.color;
      }
    }

    // Group consecutive characters with same color into spans
    const spans: { text: string; color: string }[] = [];
    let current = { text: code[0], color: charColors[0] };
    for (let i = 1; i < code.length; i++) {
      if (charColors[i] === current.color) {
        current.text += code[i];
      } else {
        spans.push(current);
        current = { text: code[i], color: charColors[i] };
      }
    }
    spans.push(current);

    return (
      <pre
        style={{
          backgroundColor: theme.background,
          color: theme.foreground,
          border: `1px solid ${theme.border}`,
          padding: "5px 10px",
          fontFamily: "Consolas, monospace",
          fontSize: 11,
          margin: "4px 0",
          borderRadius: 4,
          overflowX: "auto",
          maxHeight: 120,
          whiteSpace: "pre",
        }}
      >
        {spans.map((span, i) => (
          <span key={i} style={{ color: span.color }}>
            {span.text}
          </span>
        ))}
      </pre>
    );
  };

  return (
    <div className="panel">
      <div className="panel-section">
        <div className="panel-section-title">代码块设置</div>

        <div className="panel-row">
          <Label size="small">语言</Label>
          <Dropdown
            value={
              CODE_LANGUAGES.find((l) => l.key === language)?.label || "Python"
            }
            onOptionSelect={(_, data) =>
              setLanguage(data.optionValue || language)
            }
            size="small"
          >
            {CODE_LANGUAGES.map((l) => (
              <Option key={l.key} value={l.key}>
                {l.label}
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
          <Switch
            checked={darkTheme}
            onChange={(_, data) => setDarkTheme(data.checked)}
            label="深色主题"
          />
        </div>

        <div style={{ margin: "4px 0" }}>
          <Label size="small">代码</Label>
          <Textarea
            value={code}
            onChange={(_, data) => setCode(data.value)}
            placeholder="在此粘贴代码..."
            resize="vertical"
            style={{
              width: "100%",
              minHeight: 100,
              fontFamily: "Consolas, monospace",
              fontSize: 12,
            }}
          />
        </div>

        {code && (
          <div>
            <Label size="small">预览</Label>
            {renderPreview()}
          </div>
        )}
      </div>

      <div className="panel-actions">
        <Button appearance="primary" onClick={insertCodeBlock} size="small">
          插入代码块
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

export default CodeBlock;
