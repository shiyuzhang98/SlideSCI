import React, { useState } from "react";
import {
  Tab,
  TabList,
  SelectTabEvent,
  SelectTabData,
} from "@fluentui/react-components";
import {
  CheckboxCheckedRegular,
  ImageRegular,
  TagRegular,
  GridRegular,
  CopyRegular,
  CodeRegular,
  SettingsRegular,
} from "@fluentui/react-icons";
import SelectShapes from "./components/SelectShapes";
import ImageTitle from "./components/ImageTitle";
import ImageLabel from "./components/ImageLabel";
import ImageArrange from "./components/ImageArrange";
import FormatCopy from "./components/FormatCopy";
import CodeBlock from "./components/CodeBlock";
import SettingsPanel from "./components/SettingsPanel";

interface AppProps {
  isOfficeInitialized: boolean;
}

const TABS = [
  { value: "select", label: "选择", icon: <CheckboxCheckedRegular /> },
  { value: "title", label: "标题", icon: <ImageRegular /> },
  { value: "label", label: "标签", icon: <TagRegular /> },
  { value: "arrange", label: "排列", icon: <GridRegular /> },
  { value: "format", label: "格式", icon: <CopyRegular /> },
  { value: "code", label: "代码", icon: <CodeRegular /> },
  { value: "settings", label: "设置", icon: <SettingsRegular /> },
];

const App: React.FC<AppProps> = ({ isOfficeInitialized }) => {
  const [activeTab, setActiveTab] = useState("select");

  const onTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
    setActiveTab(data.value as string);
  };

  if (!isOfficeInitialized) {
    return (
      <div className="app-container">
        <div className="loading">
          <p>正在连接 PowerPoint...</p>
          <p className="loading-hint">请确保在 PowerPoint 中打开此插件</p>
        </div>
      </div>
    );
  }

  const renderPanel = () => {
    switch (activeTab) {
      case "select":
        return <SelectShapes />;
      case "title":
        return <ImageTitle />;
      case "label":
        return <ImageLabel />;
      case "arrange":
        return <ImageArrange />;
      case "format":
        return <FormatCopy />;
      case "code":
        return <CodeBlock />;
      case "settings":
        return <SettingsPanel />;
      default:
        return <ImageTitle />;
    }
  };

  return (
    <div className="app-container">
      <div className="tab-bar">
        <TabList
          selectedValue={activeTab}
          onTabSelect={onTabSelect}
          size="small"
          appearance="subtle"
        >
          {TABS.map((tab) => (
            <Tab key={tab.value} value={tab.value} icon={tab.icon}>
              {tab.label}
            </Tab>
          ))}
        </TabList>
      </div>
      <div className="panel-content">{renderPanel()}</div>
    </div>
  );
};

export default App;
