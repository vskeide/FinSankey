/* ---------------------------------------------------------------
 * Taskpane.tsx  –  Main add-in UI using Fluent UI v9
 * --------------------------------------------------------------- */
/* global Excel */
import React, { useState, useRef, useCallback } from "react";
import {
    Button,
    Select,
    Label,
    MessageBar,
    MessageBarBody,
    Spinner,
    Tooltip,
    Checkbox,
    tokens,
} from "@fluentui/react-components";
import {
    DataBarVertical24Regular,
    ArrowDownload24Regular,
    ArrowClockwise24Regular,
} from "@fluentui/react-icons";
import { SankeyChart, SankeyChartHandle } from "./SankeyChart";
import { parseExcelData, GraphData, ScaleMode, DecimalMode } from "./dataParser";

export const Taskpane: React.FC = () => {
    const [graphData, setGraphData] = useState<GraphData | null>(null);
    const [scale, setScale] = useState<ScaleMode>("B");
    const [showValues, setShowValues] = useState<boolean>(true);
    const [showYoy, setShowYoy] = useState<boolean>(false);
    const [decimals, setDecimals] = useState<DecimalMode>(1);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [exporting, setExporting] = useState(false);
    const chartRef = useRef<SankeyChartHandle>(null);

    // ── Load from Excel selection ──────────────────────────────────
    const loadData = useCallback(async () => {
        setLoading(true);
        setError(null);
        try {
            await Excel.run(async (ctx) => {
                const range = ctx.workbook.getSelectedRange();
                range.load("values");
                await ctx.sync();
                const parsed = parseExcelData(
                    range.values as (string | number | boolean)[][]
                );
                setGraphData(parsed);
            });
        } catch (e: unknown) {
            setError((e as Error).message ?? "Failed to read Excel selection.");
        } finally {
            setLoading(false);
        }
    }, []);

    // ── Export SVG → PNG → Excel image ────────────────────────────
    const exportToSheet = useCallback(async () => {
        const svgEl = chartRef.current?.getSvgElement();
        if (!svgEl) return;
        setExporting(true);
        try {
            const w = svgEl.width.baseVal.value;
            const h = svgEl.height.baseVal.value;
            const DPR = 2; // 2× resolution

            // Inline all styles so the off-screen canvas render is faithful
            const serializer = new XMLSerializer();
            const svgStr = serializer.serializeToString(svgEl);
            const blob = new Blob([svgStr], { type: "image/svg+xml;charset=utf-8" });
            const url = URL.createObjectURL(blob);

            const img = await new Promise<HTMLImageElement>((res, rej) => {
                const i = new Image();
                i.onload = () => res(i);
                i.onerror = rej;
                i.src = url;
            });
            URL.revokeObjectURL(url);

            const canvas = document.createElement("canvas");
            canvas.width = w * DPR;
            canvas.height = h * DPR;
            const ctx2d = canvas.getContext("2d")!;
            ctx2d.fillStyle = "#ffffff";
            ctx2d.fillRect(0, 0, canvas.width, canvas.height);
            ctx2d.scale(DPR, DPR);
            ctx2d.drawImage(img, 0, 0, w, h);

            const pngBase64 = canvas.toDataURL("image/png").split(",")[1];

            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                const shape = sheet.shapes.addImage(pngBase64);
                shape.name = "FinSankey";
                shape.width = w;
                shape.top = 20;
                shape.left = 20;
                await ctx.sync();
            });
        } catch (e: unknown) {
            setError("Export failed: " + (e as Error).message);
        } finally {
            setExporting(false);
        }
    }, []);

    // ── Dimensions (responsive to taskpane) ───────────────────────
    const CHART_W = 860;
    const CHART_H = 800;

    return (
        <div style={styles.shell}>
            {/* ── Header ─────────────────────────────────────────── */}
            <div style={styles.header}>
                <DataBarVertical24Regular style={{ color: tokens.colorBrandForeground1 }} />
                <span style={styles.title}>FinSankey</span>
            </div>

            {/* ── Toolbar ────────────────────────────────────────── */}
            <div style={styles.toolbar}>
                <Tooltip content="Select a 4-column range in Excel first (Source | Target | Amount current | Amount last year)" relationship="description">
                    <Button
                        appearance="primary"
                        icon={<ArrowClockwise24Regular />}
                        onClick={loadData}
                        disabled={loading}
                    >
                        {loading ? <Spinner size="tiny" /> : "Load Selection"}
                    </Button>
                </Tooltip>

                <div style={styles.scaleBox}>
                    <Label htmlFor="scale-sel" style={{ fontSize: 12 }}>Scale</Label>
                    <Select
                        id="scale-sel"
                        size="small"
                        value={scale}
                        onChange={(_, d) => setScale(d.value as ScaleMode)}
                    >
                        <option value="Raw">Raw</option>
                        <option value="K">K</option>
                        <option value="M">M</option>
                        <option value="B">B</option>
                    </Select>
                </div>

                <div style={styles.scaleBox}>
                    <Label style={{ fontSize: 12 }}>Show</Label>
                    <Checkbox
                        label="Values"
                        checked={showValues}
                        onChange={(_, d) => setShowValues(!!d.checked)}
                    />
                    <Checkbox
                        label="YoY"
                        checked={showYoy}
                        onChange={(_, d) => setShowYoy(!!d.checked)}
                    />
                </div>

                <div style={styles.scaleBox}>
                    <Label htmlFor="decimals-sel" style={{ fontSize: 12 }}>Decimals</Label>
                    <Select
                        id="decimals-sel"
                        size="small"
                        value={String(decimals)}
                        onChange={(_, d) => setDecimals(Number(d.value) as DecimalMode)}
                    >
                        <option value="0">0</option>
                        <option value="1">1</option>
                        <option value="2">2</option>
                    </Select>
                </div>

                <Button
                    appearance="outline"
                    icon={<ArrowDownload24Regular />}
                    onClick={exportToSheet}
                    disabled={!graphData || exporting}
                >
                    {exporting ? <Spinner size="tiny" /> : "Export to Sheet"}
                </Button>
            </div>

            {/* ── Error banner ───────────────────────────────────── */}
            {error && (
                <MessageBar intent="error" style={{ margin: "8px 12px" }}>
                    <MessageBarBody>{error}</MessageBarBody>
                </MessageBar>
            )}

            {/* ── Chart or placeholder ───────────────────────────── */}
            <div style={styles.chartArea}>
                {graphData ? (
                    <div style={{ overflowX: "auto", overflowY: "auto" }}>
                        <SankeyChart
                            ref={chartRef}
                            data={graphData}
                            scale={scale}
                            showValues={showValues}
                            showYoy={showYoy}
                            decimals={decimals}
                            width={CHART_W}
                            height={CHART_H}
                        />
                    </div>
                ) : (
                    <div style={styles.placeholder}>
                        <DataBarVertical24Regular style={{ fontSize: 48, color: "#ccc" }} />
                        <p style={{ color: "#888", marginTop: 12, textAlign: "center" }}>
                            Select a range in Excel with columns:<br />
                            <strong>Source | Target | Amount current | Amount last year</strong><br />
                            then click <strong>Load Selection</strong>.
                        </p>
                    </div>
                )}
            </div>
        </div>
    );
};

// ── Inline styles ──────────────────────────────────────────────
const styles: Record<string, React.CSSProperties> = {
    shell: {
        display: "flex",
        flexDirection: "column",
        height: "100vh",
        fontFamily: "Inter, 'Segoe UI', sans-serif",
        background: "#f9fafb",
    },
    header: {
        display: "flex",
        alignItems: "center",
        gap: 8,
        padding: "10px 14px 6px",
        borderBottom: "1px solid #e0e0e0",
        background: "#fff",
    },
    title: {
        fontWeight: 700,
        fontSize: 16,
        color: "#1a1a1a",
        letterSpacing: "-0.3px",
    },
    toolbar: {
        display: "flex",
        alignItems: "center",
        gap: 10,
        padding: "8px 12px",
        background: "#fff",
        borderBottom: "1px solid #eee",
        flexWrap: "wrap",
    },
    scaleBox: {
        display: "flex",
        alignItems: "center",
        gap: 6,
    },
    chartArea: {
        flex: 1,
        padding: "12px 6px",
        overflow: "auto",
    },
    placeholder: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        height: "100%",
        minHeight: 300,
    },
};
