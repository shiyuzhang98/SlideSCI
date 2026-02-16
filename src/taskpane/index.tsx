import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./App";
import "./styles/taskpane.css";

let isOfficeInitialized = false;

const render = () => {
  const container = document.getElementById("root");
  if (!container) return;
  const root = createRoot(container);
  root.render(
    <FluentProvider theme={webLightTheme}>
      <App isOfficeInitialized={isOfficeInitialized} />
    </FluentProvider>
  );
};

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    isOfficeInitialized = true;
  }
  render();
});
