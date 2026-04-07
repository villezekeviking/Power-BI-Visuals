"use strict";

import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

import { VisualSettings } from "./settings";

// ---------------------------------------------------------------------------
// Data model
// ---------------------------------------------------------------------------

/** A single event node in the process flow. */
export interface ProcessNode {
    id: string;
    /** Human-readable display name (falls back to id when absent). */
    name: string;
    /** Total occurrence count aggregated across all rows with this event id. */
    count: number;
    /** Horizontal layer assigned by BFS from root nodes. */
    layer: number;
    /** Computed X position (SVG coordinate). */
    posX: number;
    /** Computed Y position (SVG coordinate). */
    posY: number;
}

/** A directed transition edge between two event nodes. */
export interface ProcessEdge {
    sourceId: string;
    targetId: string;
    /** Total transition count. */
    count: number;
    /** Average milliseconds from source event to target event (null = no data). */
    avgDuration: number | null;
    /**
     * Set to true by computeLayout when this edge creates a cycle (back-edge
     * in the DFS spanning tree).  Back-edges are rendered as curved arcs so
     * that loops in the process flow are visually distinct from forward transitions.
     */
    isBackEdge?: boolean;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const SVG_NS = "http://www.w3.org/2000/svg";

function createSvgEl<T extends SVGElement>(tag: string): T {
    return document.createElementNS(SVG_NS, tag) as T;
}

function truncateLabel(text: string, diameter: number, fontSize: number): string {
    // Approx 0.6× font-size per character (standard proportional-font estimate)
    const charWidth = fontSize * 0.6;
    const maxChars = Math.max(1, Math.floor(diameter / charWidth));
    return text.length > maxChars ? text.substring(0, maxChars - 1) + "…" : text;
}

function formatDuration(ms: number): string {
    if (ms < 1000) return `${ms.toFixed(0)} ms`;
    if (ms < 60_000) return `${(ms / 1000).toFixed(1)} s`;
    if (ms < 3_600_000) return `${(ms / 60_000).toFixed(1)} min`;
    if (ms < 86_400_000) return `${(ms / 3_600_000).toFixed(1)} h`;
    return `${(ms / 86_400_000).toFixed(1)} d`;
}

// ---------------------------------------------------------------------------
// Visual class
// ---------------------------------------------------------------------------

export class Visual implements IVisual {
    private readonly host: IVisualHost;
    private readonly svg: SVGSVGElement;

    private settings: VisualSettings = new VisualSettings();

    private readonly MARGIN = { top: 60, right: 60, bottom: 60, left: 60 };

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;

        this.svg = createSvgEl<SVGSVGElement>("svg");
        this.svg.classList.add("processMappingTree");
        this.svg.style.cssText = "width:100%;height:100%;overflow:hidden;";
        options.element.appendChild(this.svg);
    }

    // ── Public API ───────────────────────────────────────────────────────────

    public update(options: VisualUpdateOptions): void {
        const dataView: DataView | undefined = options.dataViews?.[0];
        if (!dataView?.table) {
            this.renderEmpty(options.viewport.width, options.viewport.height, "Bind at least an 'Event ID' column to get started.");
            return;
        }

        this.settings = VisualSettings.parse(dataView);

        const { nodes, edges } = this.parseData(dataView);

        if (nodes.size === 0) {
            this.renderEmpty(options.viewport.width, options.viewport.height, "No event data found. Ensure the 'Event ID' column has values.");
            return;
        }

        this.computeLayout(nodes, edges, options.viewport.width, options.viewport.height);
        this.render(nodes, edges, options.viewport.width, options.viewport.height);
    }

    // ── Data parsing ─────────────────────────────────────────────────────────

    /**
     * Convert the Power BI table DataView into a graph of ProcessNode /
     * ProcessEdge objects.
     *
     * Expected columns (matched by data-role name):
     *   eventId      – required, text/numeric
     *   nextStepId   – optional, text/numeric  (null ⇒ terminal node)
     *   eventName    – optional, text           (display name)
     *   count        – optional, numeric        (occurrence count per row)
     *   avgDuration  – optional, numeric (ms)   (avg time to next step; rows are
     *                  combined using a count-weighted average)
     */
    public parseData(dataView: DataView): { nodes: Map<string, ProcessNode>; edges: ProcessEdge[] } {
        const table = dataView.table!;
        const columns = table.columns;

        let eventIdIdx = -1;
        let nextStepIdIdx = -1;
        let eventNameIdx = -1;
        let countIdx = -1;
        let avgDurationIdx = -1;

        for (let i = 0; i < columns.length; i++) {
            const roles = columns[i].roles ?? {};
            if (roles["eventId"]) eventIdIdx = i;
            if (roles["nextStepId"]) nextStepIdIdx = i;
            if (roles["eventName"]) eventNameIdx = i;
            if (roles["count"]) countIdx = i;
            if (roles["avgDuration"]) avgDurationIdx = i;
        }

        if (eventIdIdx === -1) {
            return { nodes: new Map(), edges: [] };
        }

        const nodes = new Map<string, ProcessNode>();
        const edgeMap = new Map<string, ProcessEdge>();
        // Track weighted duration sums to compute a proper average across rows
        const durationWeightedSum = new Map<string, number>();
        const durationTotalWeight = new Map<string, number>();

        for (const row of table.rows ?? []) {
            const rawId = row[eventIdIdx];
            if (rawId == null || rawId === "") continue;

            const eventId = String(rawId);
            const nextStepId = nextStepIdIdx >= 0 && row[nextStepIdIdx] != null ? String(row[nextStepIdIdx]) : null;
            const eventName = eventNameIdx >= 0 && row[eventNameIdx] != null ? String(row[eventNameIdx]) : eventId;
            const count = countIdx >= 0 && row[countIdx] != null ? Number(row[countIdx]) || 0 : 1;
            const avgDuration = avgDurationIdx >= 0 && row[avgDurationIdx] != null ? Number(row[avgDurationIdx]) : null;

            // ── Upsert source node ──────────────────────────────────────────
            if (!nodes.has(eventId)) {
                nodes.set(eventId, { id: eventId, name: eventName, count: 0, layer: -1, posX: 0, posY: 0 });
            }
            const node = nodes.get(eventId)!;
            node.count += count;
            // Prefer an explicit eventName over the id fallback
            if (eventNameIdx >= 0 && row[eventNameIdx] != null) {
                node.name = eventName;
            }

            // ── Upsert target node & edge ───────────────────────────────────
            if (nextStepId !== null && nextStepId !== eventId) {
                if (!nodes.has(nextStepId)) {
                    nodes.set(nextStepId, { id: nextStepId, name: nextStepId, count: 0, layer: -1, posX: 0, posY: 0 });
                }

                const edgeKey = `${eventId}|||${nextStepId}`;
                if (!edgeMap.has(edgeKey)) {
                    edgeMap.set(edgeKey, { sourceId: eventId, targetId: nextStepId, count: 0, avgDuration: null });
                }
                const edge = edgeMap.get(edgeKey)!;
                edge.count += count;
                // Accumulate a count-weighted sum to compute a proper average
                // across rows that each represent `count` occurrences of this transition.
                if (avgDuration != null) {
                    durationWeightedSum.set(edgeKey, (durationWeightedSum.get(edgeKey) ?? 0) + avgDuration * count);
                    durationTotalWeight.set(edgeKey, (durationTotalWeight.get(edgeKey) ?? 0) + count);
                    edge.avgDuration = durationWeightedSum.get(edgeKey)! / durationTotalWeight.get(edgeKey)!;
                }
            }
        }

        return { nodes, edges: Array.from(edgeMap.values()) };
    }

    // ── Layout ───────────────────────────────────────────────────────────────

    /**
     * Assign (posX, posY) to each node using a left-to-right layered layout.
     *
     * Algorithm:
     *   1. Build adjacency list.
     *   2. Detect back-edges (cycle-creating edges) via iterative DFS.
     *      Back-edges are flagged on the edge object (isBackEdge = true).
     *   3. Identify root nodes – nodes with no incoming *forward* edge.
     *   4. BFS from roots using only forward edges; DAG merge nodes receive
     *      the maximum (longest-path) layer across all incoming paths.
     *   5. Nodes still at layer=-1 are isolated; place them on layer 0.
     *   6. Within each layer, spread nodes evenly on the vertical axis.
     */
    public computeLayout(
        nodes: Map<string, ProcessNode>,
        edges: ProcessEdge[],
        viewWidth: number,
        viewHeight: number
    ): void {
        // 1. Build adjacency list
        const adjacency = new Map<string, string[]>();
        edges.forEach(e => {
            if (!adjacency.has(e.sourceId)) adjacency.set(e.sourceId, []);
            adjacency.get(e.sourceId)!.push(e.targetId);
        });

        // 2. Detect back-edges via iterative DFS (gray/black colouring)
        //    A back-edge is an edge whose target is still "gray"
        //    (currently on the DFS stack), meaning it leads to an ancestor.
        const backEdgeKeys = new Set<string>();
        {
            type DfsFrame = { id: string; neighbors: string[]; idx: number };
            // 0 = unvisited, 1 = in current DFS path, 2 = finished
            const color = new Map<string, number>();

            const dfsFrom = (startId: string): void => {
                const stack: DfsFrame[] = [];
                color.set(startId, 1);
                stack.push({ id: startId, neighbors: adjacency.get(startId) ?? [], idx: 0 });

                while (stack.length > 0) {
                    const frame = stack[stack.length - 1];
                    if (frame.idx < frame.neighbors.length) {
                        const nbrId = frame.neighbors[frame.idx++];
                        const c = color.get(nbrId) ?? 0;
                        if (c === 1) {
                            // Target is on the current DFS stack → back-edge (cycle)
                            backEdgeKeys.add(`${frame.id}|||${nbrId}`);
                        } else if (c === 0) {
                            color.set(nbrId, 1);
                            stack.push({ id: nbrId, neighbors: adjacency.get(nbrId) ?? [], idx: 0 });
                        }
                    } else {
                        color.set(frame.id, 2);
                        stack.pop();
                    }
                }
            };

            nodes.forEach((_, id) => {
                if ((color.get(id) ?? 0) === 0) dfsFrom(id);
            });
        }

        // Mark back-edges on the edge objects so the renderer can draw them
        // as curved arcs instead of straight lines.
        edges.forEach(e => {
            e.isBackEdge = backEdgeKeys.has(`${e.sourceId}|||${e.targetId}`);
        });

        // 3. Find roots: nodes with no incoming *forward* (non-back) edge
        const hasIncomingForward = new Set<string>();
        edges.forEach(e => { if (!e.isBackEdge) hasIncomingForward.add(e.targetId); });
        const roots = Array.from(nodes.keys()).filter(id => !hasIncomingForward.has(id));

        if (roots.length === 0) {
            // Fully cyclic graph – pick the node with the most outgoing edges as root
            const outDegree = new Map<string, number>();
            edges.forEach(e => outDegree.set(e.sourceId, (outDegree.get(e.sourceId) ?? 0) + 1));
            const best = Array.from(nodes.keys()).sort((a, b) => (outDegree.get(b) ?? 0) - (outDegree.get(a) ?? 0))[0];
            roots.push(best);
        }

        // 4. BFS layer assignment (forward edges only)
        //    When a node is reached via multiple paths we keep the *deepest*
        //    (longest-path) layer so merge nodes are placed correctly.
        const visited = new Set<string>();
        const queue: string[] = [];

        roots.forEach(id => {
            nodes.get(id)!.layer = 0;
            visited.add(id);
            queue.push(id);
        });

        while (queue.length > 0) {
            const currentId = queue.shift()!;
            const currentLayer = nodes.get(currentId)!.layer;

            for (const targetId of adjacency.get(currentId) ?? []) {
                // Never follow back-edges during layer assignment
                if (backEdgeKeys.has(`${currentId}|||${targetId}`)) continue;

                const target = nodes.get(targetId);
                if (!target) continue;

                if (!visited.has(targetId)) {
                    target.layer = currentLayer + 1;
                    visited.add(targetId);
                    queue.push(targetId);
                } else if (target.layer <= currentLayer) {
                    // DAG merge node reached via a longer path – push it right
                    target.layer = currentLayer + 1;
                }
            }
        }

        // 5. Fix isolated nodes
        nodes.forEach(node => { if (node.layer === -1) node.layer = 0; });

        // 4. Group by layer and assign positions
        const layerGroups = new Map<number, ProcessNode[]>();
        nodes.forEach(node => {
            if (!layerGroups.has(node.layer)) layerGroups.set(node.layer, []);
            layerGroups.get(node.layer)!.push(node);
        });

        const sortedLayers = Array.from(layerGroups.keys()).sort((a, b) => a - b);
        const numLayers = sortedLayers.length;

        const contentW = viewWidth - this.MARGIN.left - this.MARGIN.right;
        const contentH = viewHeight - this.MARGIN.top - this.MARGIN.bottom;

        const layerXForIndex = (idx: number): number => {
            if (numLayers === 1) return this.MARGIN.left + contentW / 2;
            return this.MARGIN.left + (idx / (numLayers - 1)) * contentW;
        };

        sortedLayers.forEach((layer, layerIdx) => {
            const layerNodes = layerGroups.get(layer)!;
            const n = layerNodes.length;
            const x = layerXForIndex(layerIdx);

            layerNodes.forEach((node, i) => {
                node.posX = x;
                node.posY = n === 1
                    ? this.MARGIN.top + contentH / 2
                    : this.MARGIN.top + (i / (n - 1)) * contentH;
            });
        });
    }

    // ── Rendering ────────────────────────────────────────────────────────────

    private render(
        nodes: Map<string, ProcessNode>,
        edges: ProcessEdge[],
        width: number,
        height: number
    ): void {
        const { nodeColor, nodeRadius, showCount } = this.settings.nodeSettings;
        const { edgeColor, showTimeMetrics } = this.settings.edgeSettings;
        const { fontSize, fontColor } = this.settings.labelSettings;

        this.clearSvg(width, height);

        // ── Defs (arrow markers) ────────────────────────────────────────────
        const defs = createSvgEl<SVGDefsElement>("defs");

        const makeMarker = (id: string, color: string, refX: number): SVGMarkerElement => {
            const m = createSvgEl<SVGMarkerElement>("marker");
            m.setAttribute("id", id);
            m.setAttribute("viewBox", "0 -5 10 10");
            m.setAttribute("refX", String(refX));
            m.setAttribute("refY", "0");
            m.setAttribute("markerWidth", "6");
            m.setAttribute("markerHeight", "6");
            m.setAttribute("orient", "auto");
            const p = createSvgEl<SVGPathElement>("path");
            p.setAttribute("d", "M0,-5L10,0L0,5");
            p.setAttribute("fill", color);
            m.appendChild(p);
            return m;
        };

        defs.appendChild(makeMarker("pmt-arrow", edgeColor, nodeRadius + 10));
        defs.appendChild(makeMarker("pmt-arrow-back", edgeColor, nodeRadius + 10));
        this.svg.appendChild(defs);

        // ── Title ───────────────────────────────────────────────────────────
        const title = createSvgEl<SVGTextElement>("text");
        title.setAttribute("x", String(width / 2));
        title.setAttribute("y", "24");
        title.setAttribute("text-anchor", "middle");
        title.setAttribute("font-size", "14");
        title.setAttribute("font-weight", "bold");
        title.setAttribute("fill", "#333");
        title.textContent = "Process Mapping Tree";
        this.svg.appendChild(title);

        // ── Edges ───────────────────────────────────────────────────────────
        const edgeGroup = createSvgEl<SVGGElement>("g");
        edgeGroup.classList.add("pmt-edges");

        edges.forEach(edge => {
            const source = nodes.get(edge.sourceId);
            const target = nodes.get(edge.targetId);
            if (!source || !target) return;

            // Stroke width scales (logarithmically) with transition count
            const strokeWidth = Math.max(1.5, Math.log2(edge.count + 1) * 1.2);

            if (edge.isBackEdge) {
                // ── Back-edge: curved arc going above the flow ──────────────
                // Use a quadratic bezier whose control point arcs over the nodes.
                const arcHeight = Math.max(80, Math.abs(source.posX - target.posX) * 0.45 + 40);
                const cpX = (source.posX + target.posX) / 2;
                const cpY = Math.min(source.posY, target.posY) - arcHeight;

                const curvePath = createSvgEl<SVGPathElement>("path");
                curvePath.setAttribute("d",
                    `M${source.posX},${source.posY} Q${cpX},${cpY} ${target.posX},${target.posY}`);
                curvePath.setAttribute("fill", "none");
                curvePath.setAttribute("stroke", edgeColor);
                curvePath.setAttribute("stroke-width", String(strokeWidth));
                curvePath.setAttribute("stroke-dasharray", "7,4");
                curvePath.setAttribute("marker-end", "url(#pmt-arrow-back)");
                curvePath.setAttribute("class", "pmt-back-edge");
                edgeGroup.appendChild(curvePath);

                // Labels near the arc apex
                const lblX = cpX;
                const lblY = cpY - 4;

                const countLbl = createSvgEl<SVGTextElement>("text");
                countLbl.setAttribute("x", String(lblX));
                countLbl.setAttribute("y", String(lblY));
                countLbl.setAttribute("text-anchor", "middle");
                countLbl.setAttribute("font-size", String(Math.max(9, fontSize - 1)));
                countLbl.setAttribute("fill", "#555");
                countLbl.setAttribute("class", "pmt-edge-label");
                countLbl.textContent = `×${edge.count}`;
                edgeGroup.appendChild(countLbl);

                if (showTimeMetrics && edge.avgDuration != null) {
                    const durLbl = createSvgEl<SVGTextElement>("text");
                    durLbl.setAttribute("x", String(lblX));
                    durLbl.setAttribute("y", String(lblY - fontSize - 2));
                    durLbl.setAttribute("text-anchor", "middle");
                    durLbl.setAttribute("font-size", String(Math.max(9, fontSize - 1)));
                    durLbl.setAttribute("fill", "#777");
                    durLbl.setAttribute("class", "pmt-edge-label");
                    durLbl.textContent = `⌀ ${formatDuration(edge.avgDuration)}`;
                    edgeGroup.appendChild(durLbl);
                }
            } else {
                // ── Forward edge: straight line ─────────────────────────────
                const line = createSvgEl<SVGLineElement>("line");
                line.setAttribute("x1", String(source.posX));
                line.setAttribute("y1", String(source.posY));
                line.setAttribute("x2", String(target.posX));
                line.setAttribute("y2", String(target.posY));
                line.setAttribute("stroke", edgeColor);
                line.setAttribute("stroke-width", String(strokeWidth));
                line.setAttribute("marker-end", "url(#pmt-arrow)");
                edgeGroup.appendChild(line);

                // Mid-point for labels
                const midX = (source.posX + target.posX) / 2;
                const midY = (source.posY + target.posY) / 2;

                // Transition count label
                const countLbl = createSvgEl<SVGTextElement>("text");
                countLbl.setAttribute("x", String(midX));
                countLbl.setAttribute("y", String(midY - (showTimeMetrics && edge.avgDuration != null ? 8 : 0)));
                countLbl.setAttribute("text-anchor", "middle");
                countLbl.setAttribute("font-size", String(Math.max(9, fontSize - 1)));
                countLbl.setAttribute("fill", "#555");
                countLbl.setAttribute("class", "pmt-edge-label");
                countLbl.textContent = `×${edge.count}`;
                edgeGroup.appendChild(countLbl);

                // Avg duration label
                if (showTimeMetrics && edge.avgDuration != null) {
                    const durLbl = createSvgEl<SVGTextElement>("text");
                    durLbl.setAttribute("x", String(midX));
                    durLbl.setAttribute("y", String(midY + fontSize));
                    durLbl.setAttribute("text-anchor", "middle");
                    durLbl.setAttribute("font-size", String(Math.max(9, fontSize - 1)));
                    durLbl.setAttribute("fill", "#777");
                    durLbl.setAttribute("class", "pmt-edge-label");
                    durLbl.textContent = `⌀ ${formatDuration(edge.avgDuration)}`;
                    edgeGroup.appendChild(durLbl);
                }
            }
        });

        this.svg.appendChild(edgeGroup);

        // ── Nodes ───────────────────────────────────────────────────────────
        const nodeGroup = createSvgEl<SVGGElement>("g");
        nodeGroup.classList.add("pmt-nodes");

        nodes.forEach(node => {
            const g = createSvgEl<SVGGElement>("g");
            g.setAttribute("transform", `translate(${node.posX},${node.posY})`);
            g.classList.add("pmt-node");

            // Circle
            const circle = createSvgEl<SVGCircleElement>("circle");
            circle.setAttribute("r", String(nodeRadius));
            circle.setAttribute("fill", nodeColor);
            circle.setAttribute("stroke", "#1a1a2e");
            circle.setAttribute("stroke-width", "2.5");
            g.appendChild(circle);

            // Event name label (inside circle)
            const nameLbl = createSvgEl<SVGTextElement>("text");
            nameLbl.setAttribute("text-anchor", "middle");
            nameLbl.setAttribute("dominant-baseline", "middle");
            nameLbl.setAttribute("font-size", String(fontSize));
            nameLbl.setAttribute("font-weight", "600");
            nameLbl.setAttribute("fill", fontColor);
            nameLbl.textContent = truncateLabel(node.name, nodeRadius * 2, fontSize);
            g.appendChild(nameLbl);

            // Count label below circle
            if (showCount && node.count > 0) {
                const cntLbl = createSvgEl<SVGTextElement>("text");
                cntLbl.setAttribute("text-anchor", "middle");
                cntLbl.setAttribute("y", String(nodeRadius + fontSize + 4));
                cntLbl.setAttribute("font-size", String(Math.max(9, fontSize - 1)));
                cntLbl.setAttribute("fill", "#333");
                cntLbl.textContent = `n = ${node.count}`;
                g.appendChild(cntLbl);
            }

            nodeGroup.appendChild(g);
        });

        this.svg.appendChild(nodeGroup);

        // ── Legend ──────────────────────────────────────────────────────────
        this.renderLegend(width, height, edgeColor, nodeColor, showTimeMetrics);
    }

    /** Render a small legend in the bottom-right corner. */
    private renderLegend(
        width: number,
        height: number,
        edgeColor: string,
        nodeColor: string,
        showTimeMetrics: boolean
    ): void {
        const legend = createSvgEl<SVGGElement>("g");
        legend.classList.add("pmt-legend");

        const items: string[] = [
            "● Node = event type",
            "— Edge = transition (×count)",
            "- - Dashed = loop / back-edge",
        ];
        if (showTimeMetrics) items.push("⌀ Avg time between steps");

        const lineH = 16;
        const legendW = 200;
        const legendH = items.length * lineH + 16;
        const lx = width - legendW - 10;
        const ly = height - legendH - 10;

        const bg = createSvgEl<SVGRectElement>("rect");
        bg.setAttribute("x", String(lx - 6));
        bg.setAttribute("y", String(ly - 6));
        bg.setAttribute("width", String(legendW));
        bg.setAttribute("height", String(legendH));
        bg.setAttribute("rx", "4");
        bg.setAttribute("fill", "rgba(255,255,255,0.85)");
        bg.setAttribute("stroke", "#ccc");
        bg.setAttribute("stroke-width", "1");
        legend.appendChild(bg);

        items.forEach((item, i) => {
            const t = createSvgEl<SVGTextElement>("text");
            t.setAttribute("x", String(lx));
            t.setAttribute("y", String(ly + i * lineH + lineH / 2));
            t.setAttribute("dominant-baseline", "middle");
            t.setAttribute("font-size", "10");
            t.setAttribute("fill", "#555");
            t.textContent = item;
            legend.appendChild(t);
        });

        this.svg.appendChild(legend);
    }

    private renderEmpty(width: number, height: number, message: string): void {
        this.clearSvg(width, height);

        const text = createSvgEl<SVGTextElement>("text");
        text.setAttribute("x", String(width / 2));
        text.setAttribute("y", String(height / 2));
        text.setAttribute("text-anchor", "middle");
        text.setAttribute("dominant-baseline", "middle");
        text.setAttribute("font-size", "13");
        text.setAttribute("fill", "#888");
        text.textContent = message;
        this.svg.appendChild(text);
    }

    private clearSvg(width: number, height: number): void {
        while (this.svg.firstChild) this.svg.removeChild(this.svg.firstChild);
        this.svg.setAttribute("width", String(width));
        this.svg.setAttribute("height", String(height));
    }
}
