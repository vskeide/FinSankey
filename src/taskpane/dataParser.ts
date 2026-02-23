/* ---------------------------------------------------------------
 * dataParser.ts
 * Parses a 4-column Excel range and prepares data for d3-sankey.
 * Columns expected: Source | Target | Amount current | Amount last year
 * --------------------------------------------------------------- */

export type ScaleMode = "Raw" | "K" | "M" | "B";
export type DecimalMode = 0 | 1 | 2;

// ── Color scheme matching SankeyArt reference ──────────────────
const COLORS = {
    profit: { fill: "#4CAF50", accent: "#2E7D32", link: "rgba(76,175,80,0.5)" },
    cost: { fill: "#E91E63", accent: "#880E4F", link: "rgba(233,30,99,0.5)" },
    neutral: { fill: "#9E9E9E", accent: "#424242", link: "rgba(158,158,158,0.45)" },
};

// ── Public types ────────────────────────────────────────────────
export interface NodeMeta {
    id: string;
    name: string;
    totalValue: number;
    totalPrior: number;
    yoyPct: number | null;
    fillColor: string;
    accentColor: string;
}

export interface LinkMeta {
    sourceId: string;
    targetId: string;
    value: number;
    prior: number;
    yoyPct: number | null;
    linkColor: string;
}

export interface GraphData {
    nodeMeta: Map<string, NodeMeta>;
    linkMeta: LinkMeta[];
    /** Raw input for d3-sankey */
    sankeyNodes: { id: string }[];
    sankeyLinks: { source: string; target: string; value: number }[];
}

// ── Helpers ─────────────────────────────────────────────────────
function parseNum(raw: string | number | boolean): number {
    if (typeof raw === "number") return raw;
    // Accept both comma and period as decimal separator
    return parseFloat(String(raw).trim().replace(/\s/g, "").replace(",", ".")) || 0;
}

// ── Main parser ─────────────────────────────────────────────────
export function parseExcelData(
    rawValues: (string | number | boolean)[][]
): GraphData {
    if (!rawValues || rawValues.length < 2) {
        throw new Error(
            "Select at least 2 rows (header + 1 data row) with columns: Source | Target | Amount current | Amount last year"
        );
    }

    // Detect header: first row col-2 is non-numeric → skip it
    const first = rawValues[0];
    const hasHeader =
        typeof first[2] === "string" &&
        isNaN(parseFloat(String(first[2]).replace(",", ".")));
    const rows = hasHeader ? rawValues.slice(1) : rawValues;

    // Accumulate totals by node id
    const nodeTotals = new Map<string, { inVal: number; outVal: number; inPrior: number; outPrior: number; rawSum: number; rawOutSum: number }>();
    const linkMeta: LinkMeta[] = [];
    const sankeyLinks: { source: string; target: string; value: number }[] = [];

    const touch = (id: string) => {
        if (!nodeTotals.has(id)) nodeTotals.set(id, { inVal: 0, outVal: 0, inPrior: 0, outPrior: 0, rawSum: 0, rawOutSum: 0 });
    };

    for (const row of rows) {
        if (!row[0] || !row[1]) continue;

        const srcId = String(row[0]).trim();
        const tgtId = String(row[1]).trim();
        const val = parseNum(row[2]);
        const prior = parseNum(row[3] ?? 0);

        touch(srcId);
        touch(tgtId);

        const absVal = Math.abs(val);
        const absPrior = Math.abs(prior);
        const yoyPct = prior !== 0 ? ((val - prior) / Math.abs(prior)) * 100 : null;

        // Accumulate to source node (outgoing flow)
        nodeTotals.get(srcId)!.outVal += absVal;
        nodeTotals.get(srcId)!.outPrior += absPrior;
        nodeTotals.get(srcId)!.rawOutSum += val;

        const cls = val < 0 ? "cost" : "profit";
        linkMeta.push({ sourceId: srcId, targetId: tgtId, value: absVal, prior: absPrior, yoyPct, linkColor: COLORS[cls].link });
        sankeyLinks.push({ source: srcId, target: tgtId, value: absVal || 0.001 });

        // Accumulate to target node (incoming flow)
        nodeTotals.get(tgtId)!.inVal += absVal;
        nodeTotals.get(tgtId)!.inPrior += absPrior;
        nodeTotals.get(tgtId)!.rawSum += val;
    }

    // Build NodeMeta map
    const nodeMeta = new Map<string, NodeMeta>();
    for (const [id, totals] of nodeTotals.entries()) {
        let cls: keyof typeof COLORS = "profit";

        // Target nodes determine class by what flows into them
        if (totals.inVal > 0) {
            cls = totals.rawSum < 0 ? "cost" : "profit";
        }
        // Root nodes determine class by what flows out of them
        else if (totals.outVal > 0) {
            cls = totals.rawOutSum < 0 ? "cost" : "profit";
        }

        const nodeVal = Math.max(totals.inVal, totals.outVal);
        const nodePrior = Math.max(totals.inPrior, totals.outPrior);

        const yoyPct = nodePrior !== 0
            ? ((nodeVal - nodePrior) / nodePrior) * 100
            : null;

        nodeMeta.set(id, {
            id,
            name: id,
            totalValue: nodeVal,
            totalPrior: nodePrior,
            yoyPct,
            fillColor: COLORS[cls].fill,
            accentColor: COLORS[cls].accent,
        });
    }

    return {
        nodeMeta,
        linkMeta,
        sankeyNodes: Array.from(nodeTotals.keys()).map((id) => ({ id })),
        sankeyLinks,
    };
}

// ── Number formatting ────────────────────────────────────────────
export function formatValue(value: number, scale: ScaleMode, decimals: number = 1): string {
    let scaled = value;
    let suffix = "";
    switch (scale) {
        case "K": scaled = value / 1_000; suffix = "K"; break;
        case "M": scaled = value / 1_000_000; suffix = "M"; break;
        case "B": scaled = value / 1_000_000_000; suffix = "B"; break;
    }
    // Comma as decimal delimiter as per spec
    return scaled.toFixed(decimals).replace(".", ",") + suffix;
}

export function formatYoy(pct: number | null, decimals: number = 0): string {
    if (pct === null) return "";
    const sign = pct >= 0 ? "+" : "";
    return `${sign}${pct.toFixed(decimals).replace(".", ",")}% Y/Y`;
}
