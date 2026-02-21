/* ---------------------------------------------------------------
 * SankeyChart.tsx
 * D3-powered Sankey renderer with:
 *   - Colored node bars (accent header strip)
 *   - Gradient links coloured by source→target
 *   - Three-line labels: Name / Value / % Y/Y
 *   - D3-drag for live node repositioning
 * --------------------------------------------------------------- */
import React, { useEffect, useRef, forwardRef, useImperativeHandle } from "react";
import * as d3 from "d3";
import {
    sankey as d3Sankey,
    sankeyLinkHorizontal,
} from "d3-sankey";
import { GraphData, ScaleMode, formatValue, formatYoy } from "./dataParser";

// ── Layout constants ────────────────────────────────────────────
const NODE_WIDTH = 16;
const NODE_PAD = 14;
const MARGIN = { top: 24, right: 200, bottom: 24, left: 200 };
const ACCENT_H = 5;   // height of the top colour bar on each node

// ── Types passed into d3-sankey ─────────────────────────────────
interface SNode { id: string;[k: string]: unknown }
interface SLink { source: string; target: string; value: number;[k: string]: unknown }

// After d3-sankey decorates them:
type LNode = SNode & {
    x0: number; x1: number; y0: number; y1: number;
    value: number; depth: number; sourceLinks: LLink[]; targetLinks: LLink[];
};
type LLink = {
    source: LNode; target: LNode;
    value: number; width: number; y0: number; y1: number;
    index: number;
};

// ── Props / ref handle ──────────────────────────────────────────
export interface SankeyChartHandle {
    getSvgElement: () => SVGSVGElement | null;
}

interface Props {
    data: GraphData;
    scale: ScaleMode;
    width?: number;
    height?: number;
}

export const SankeyChart = forwardRef<SankeyChartHandle, Props>(
    ({ data, scale, width = 860, height = 520 }, ref) => {
        const svgRef = useRef<SVGSVGElement>(null);

        useImperativeHandle(ref, () => ({
            getSvgElement: () => svgRef.current,
        }));

        useEffect(() => {
            const svg = svgRef.current;
            if (!svg || data.sankeyNodes.length === 0) return;

            const sel = d3.select(svg);
            sel.selectAll("*").remove();

            const W = width - MARGIN.left - MARGIN.right;
            const H = height - MARGIN.top - MARGIN.bottom;

            // ── Defs container ──────────────────────────────────────────
            const defs = sel.append("defs");

            // ── Main group ──────────────────────────────────────────────
            const g = sel
                .append("g")
                .attr("transform", `translate(${MARGIN.left},${MARGIN.top})`);

            // ── Sankey layout ───────────────────────────────────────────
            const layout = d3Sankey<SNode, SLink>()
                .nodeId((d) => d.id as string)
                .nodeWidth(NODE_WIDTH)
                .nodePadding(NODE_PAD)
                .extent([[0, 0], [W, H]]);

            const graph = layout({
                nodes: data.sankeyNodes.map((n) => ({ ...n })) as SNode[],
                links: data.sankeyLinks.map((l) => ({ ...l })) as SLink[],
            });

            const nodes = graph.nodes as unknown as LNode[];
            const links = graph.links as unknown as LLink[];

            // ── Gradient defs for links ─────────────────────────────────
            links.forEach((lk, i) => {
                const sM = data.nodeMeta.get(lk.source.id);
                const tM = data.nodeMeta.get(lk.target.id);
                const grad = defs
                    .append("linearGradient")
                    .attr("id", `lg-${i}`)
                    .attr("gradientUnits", "userSpaceOnUse")
                    .attr("x1", lk.source.x1)
                    .attr("x2", lk.target.x0);
                grad.append("stop").attr("offset", "0%")
                    .attr("stop-color", sM?.fillColor ?? "#9E9E9E").attr("stop-opacity", 0.55);
                grad.append("stop").attr("offset", "100%")
                    .attr("stop-color", tM?.fillColor ?? "#9E9E9E").attr("stop-opacity", 0.55);
            });

            // ── Link paths ──────────────────────────────────────────────
            const linkG = g.append("g").attr("class", "links");

            const linkPaths = linkG.selectAll<SVGPathElement, LLink>("path")
                .data(links)
                .join("path")
                .attr("d", sankeyLinkHorizontal() as never)
                .attr("stroke", (_, i) => `url(#lg-${i})`)
                .attr("stroke-width", (d) => Math.max(1, d.width))
                .attr("fill", "none")
                .attr("opacity", 0.72)
                .style("cursor", "pointer")
                .on("mouseover", function () { d3.select(this).attr("opacity", 1); })
                .on("mouseout", function () { d3.select(this).attr("opacity", 0.72); });

            // ── Node groups ─────────────────────────────────────────────
            const nodeG = g.append("g").attr("class", "nodes");

            const nodeGs = nodeG
                .selectAll<SVGGElement, LNode>("g")
                .data(nodes)
                .join("g")
                .attr("class", "node")
                .attr("transform", (d) => `translate(${d.x0},${d.y0})`)
                .style("cursor", "grab");

            // Node body
            nodeGs.append("rect")
                .attr("width", (d) => d.x1 - d.x0)
                .attr("height", (d) => Math.max(1, d.y1 - d.y0))
                .attr("fill", (d) => data.nodeMeta.get(d.id)?.fillColor ?? "#9E9E9E")
                .attr("opacity", 0.88)
                .attr("rx", 2);

            // Accent bar (top strip)
            nodeGs.append("rect")
                .attr("width", (d) => d.x1 - d.x0)
                .attr("height", ACCENT_H)
                .attr("fill", (d) => data.nodeMeta.get(d.id)?.accentColor ?? "#424242")
                .attr("rx", 2);

            // ── Labels ──────────────────────────────────────────────────
            const maxDepth = Math.max(...nodes.map((n) => n.depth));

            const labelG = nodeGs.append("g").attr("class", "label")
                .attr("transform", (d) => {
                    const nodeH = d.y1 - d.y0;
                    const onLeft = d.depth === 0;
                    const x = onLeft ? -(NODE_WIDTH / 2 + 6) : (d.x1 - d.x0) + 8;
                    const y = nodeH / 2 - 18; // vertically centre the 3-line block
                    return `translate(${x},${y})`;
                });

            const anchor = (d: LNode) => d.depth === 0 ? "end" : "start";

            // Line 1: name
            labelG.append("text")
                .attr("text-anchor", anchor)
                .attr("font-size", "11px")
                .attr("font-weight", "600")
                .attr("fill", "#222")
                .text((d) => data.nodeMeta.get(d.id)?.name ?? d.id);

            // Line 2: formatted value
            labelG.append("text")
                .attr("y", 14)
                .attr("text-anchor", anchor)
                .attr("font-size", "10px")
                .attr("fill", "#555")
                .text((d) => {
                    const m = data.nodeMeta.get(d.id);
                    return formatValue(m?.totalValue ?? d.value ?? 0, scale);
                });

            // Line 3: % Y/Y
            labelG.append("text")
                .attr("y", 27)
                .attr("text-anchor", anchor)
                .attr("font-size", "10px")
                .attr("font-weight", "500")
                .attr("fill", (d) => {
                    const pct = data.nodeMeta.get(d.id)?.yoyPct ?? null;
                    if (pct === null) return "#999";
                    return pct >= 0 ? "#2E7D32" : "#C62828";
                })
                .text((d) => formatYoy(data.nodeMeta.get(d.id)?.yoyPct ?? null));

            // ── Drag behaviour ──────────────────────────────────────────
            const drag = d3.drag<SVGGElement, LNode>()
                .on("start", function () {
                    d3.select(this).style("cursor", "grabbing").raise();
                })
                .on("drag", function (event, d) {
                    const nodeH = d.y1 - d.y0;
                    d.y0 = Math.max(0, Math.min(H - nodeH, d.y0 + event.dy));
                    d.y1 = d.y0 + nodeH;
                    d3.select(this).attr("transform", `translate(${d.x0},${d.y0})`);
                    // Re-draw all links that touch this node
                    linkPaths.filter(
                        (lk) => lk.source === (d as unknown) || lk.target === (d as unknown)
                    ).attr("d", sankeyLinkHorizontal() as never);
                })
                .on("end", function () {
                    d3.select(this).style("cursor", "grab");
                });

            nodeGs.call(drag);
        }, [data, scale, width, height]);

        return (
            <svg
                ref={svgRef}
                width={width}
                height={height}
                style={{ fontFamily: "Inter, 'Segoe UI', sans-serif", overflow: "visible" }}
            />
        );
    }
);

SankeyChart.displayName = "SankeyChart";
