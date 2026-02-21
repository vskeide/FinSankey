/* global Office */
import React from "react";
import ReactDOM from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { Taskpane } from "./Taskpane";
import "./styles.css";

Office.onReady(() => {
    const root = ReactDOM.createRoot(
        document.getElementById("root") as HTMLElement
    );
    root.render(
        <FluentProvider theme={webLightTheme}>
            <Taskpane />
        </FluentProvider>
    );
});
