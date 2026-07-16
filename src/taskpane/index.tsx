import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./App";
import "./styles/global.css";
import "./styles/markdown.css";

// Side-effect import: регистрирует все инструменты в toolRegistry ДО рендера App.
// Без этого useAgent отправляет в бэкенд пустой массив tools → у LLM нет
// инструментов для вызова. См. src/taskpane/tools/index.ts.
import "./tools";

const container = document.getElementById("root");
if (container) {
  const root = createRoot(container);
  root.render(
    <FluentProvider theme={webLightTheme}>
      <App />
    </FluentProvider>,
  );
}
