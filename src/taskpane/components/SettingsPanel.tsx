import React, { useState } from "react";
import { Button } from "@fluentui/react-components";
import { settingsService } from "../../services/settings-service";

const SettingsPanel: React.FC = () => {
  const [status, setStatus] = useState("");

  const resetSettings = () => {
    settingsService.resetToDefaults();
    setStatus("设置已重置为默认值");
  };

  return (
    <div className="panel">
      <div className="panel-section">
        <div className="panel-section-title">通用设置</div>
        <p style={{ color: "#666", fontSize: 12, marginTop: 8 }}>
          SlideSCI v1.0.0 - Office Web Add-in
        </p>
        <p style={{ color: "#666", fontSize: 12 }}>
          科研 PPT 助手，支持 macOS / Windows / Web
        </p>
      </div>

      <div className="panel-section">
        <Button appearance="secondary" onClick={resetSettings} size="small">
          重置所有设置
        </Button>
      </div>

      {status && (
        <div className="status-message success">{status}</div>
      )}

      <div className="panel-section" style={{ marginTop: 16 }}>
        <div className="panel-section-title">关于</div>
        <p style={{ color: "#666", fontSize: 12, marginTop: 4 }}>
          开发者: Achuan-2
        </p>
        <p style={{ color: "#666", fontSize: 12 }}>
          GitHub: <a href="https://github.com/Achuan-2/SlideSCI" target="_blank" rel="noreferrer">Achuan-2/SlideSCI</a>
        </p>
      </div>
    </div>
  );
};

export default SettingsPanel;
