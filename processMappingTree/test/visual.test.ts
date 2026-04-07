/**
 * Unit tests for the Process Mapping Tree visual.
 *
 * These tests exercise the two pure-logic methods – parseData and
 * computeLayout – without needing an actual Power BI host.
 */

import { Visual, ProcessNode, ProcessEdge } from "../src/visual";
import { VisualSettings } from "../src/settings";

// ---------------------------------------------------------------------------
// Minimal DataView factory helpers
// ---------------------------------------------------------------------------

type RoleMap = Record<string, boolean>;

function makeColumn(roleName: string, extra: Partial<{ displayName: string }> = {}): powerbi.DataViewMetadataColumn {
    return {
        displayName: extra.displayName ?? roleName,
        roles: { [roleName]: true },
        type: { text: true } as powerbi.ValueTypeDescriptor,
        index: 0,
    };
}

function makeTableDataView(
    columns: powerbi.DataViewMetadataColumn[],
    rows: powerbi.DataViewTableRow[]
): powerbi.DataView {
    return {
        metadata: { columns },
        table: {
            columns,
            rows,
            identity: [],
            totals: [],
        },
    } as unknown as powerbi.DataView;
}

// ---------------------------------------------------------------------------
// Helper: create a Visual instance with a stub host
// ---------------------------------------------------------------------------

function makeVisual(): Visual {
    const container = document.createElement("div");
    const host = {} as powerbi.extensibility.visual.IVisualHost;
    return new Visual({ element: container, host });
}

// ---------------------------------------------------------------------------
// parseData tests
// ---------------------------------------------------------------------------

describe("Visual.parseData", () => {
    let visual: Visual;

    beforeEach(() => {
        visual = makeVisual();
    });

    it("returns empty result when eventId column is missing", () => {
        const dv = makeTableDataView(
            [makeColumn("eventName")],
            [["Checkout"]]
        );
        const { nodes, edges } = visual.parseData(dv);
        expect(nodes.size).toBe(0);
        expect(edges.length).toBe(0);
    });

    it("parses a single event with no next step", () => {
        const dv = makeTableDataView(
            [makeColumn("eventId"), makeColumn("eventName")],
            [["A", "Step A"]]
        );
        const { nodes, edges } = visual.parseData(dv);
        expect(nodes.size).toBe(1);
        expect(nodes.get("A")?.name).toBe("Step A");
        expect(edges.length).toBe(0);
    });

    it("creates a directed edge between two events", () => {
        const dv = makeTableDataView(
            [makeColumn("eventId"), makeColumn("nextStepId"), makeColumn("eventName")],
            [
                ["A", "B", "Step A"],
                ["B", undefined as unknown as string, "Step B"],
            ]
        );
        const { nodes, edges } = visual.parseData(dv);
        expect(nodes.size).toBe(2);
        expect(edges.length).toBe(1);
        expect(edges[0].sourceId).toBe("A");
        expect(edges[0].targetId).toBe("B");
    });

    it("accumulates count across multiple rows with the same eventId", () => {
        const countCol = makeColumn("count");
        countCol.type = { numeric: true } as powerbi.ValueTypeDescriptor;

        const dv = makeTableDataView(
            [makeColumn("eventId"), countCol],
            [
                ["A", 5],
                ["A", 3],
            ]
        );
        const { nodes } = visual.parseData(dv);
        expect(nodes.get("A")?.count).toBe(8);
    });

    it("aggregates transition counts for the same (source, target) pair", () => {
        const countCol = makeColumn("count");
        countCol.type = { numeric: true } as powerbi.ValueTypeDescriptor;

        const dv = makeTableDataView(
            [makeColumn("eventId"), makeColumn("nextStepId"), countCol],
            [
                ["A", "B", 10],
                ["A", "B", 7],
            ]
        );
        const { edges } = visual.parseData(dv);
        expect(edges.length).toBe(1);
        expect(edges[0].count).toBe(17);
    });

    it("stores avgDuration on the edge", () => {
        const durCol = makeColumn("avgDuration");
        durCol.type = { numeric: true } as powerbi.ValueTypeDescriptor;

        const dv = makeTableDataView(
            [makeColumn("eventId"), makeColumn("nextStepId"), durCol],
            [["A", "B", 3000]]
        );
        const { edges } = visual.parseData(dv);
        expect(edges[0].avgDuration).toBe(3000);
    });

    it("computes a count-weighted average of avgDuration across multiple rows", () => {
        // Row 1: A→B, 100 occurrences, avg 1000 ms
        // Row 2: A→B,  50 occurrences, avg 4000 ms
        // Expected weighted avg = (100*1000 + 50*4000) / 150 = 300000/150 = 2000 ms
        const countCol = makeColumn("count");
        countCol.type = { numeric: true } as powerbi.ValueTypeDescriptor;
        const durCol = makeColumn("avgDuration");
        durCol.type = { numeric: true } as powerbi.ValueTypeDescriptor;

        const dv = makeTableDataView(
            [makeColumn("eventId"), makeColumn("nextStepId"), countCol, durCol],
            [
                ["A", "B", 100, 1000],
                ["A", "B", 50, 4000],
            ]
        );
        const { edges } = visual.parseData(dv);
        expect(edges.length).toBe(1);
        expect(edges[0].avgDuration).toBeCloseTo(2000, 5);
    });

    it("skips rows where eventId is null or empty string", () => {
        const dv = makeTableDataView(
            [makeColumn("eventId"), makeColumn("nextStepId")],
            [
                [undefined as unknown as string, "B"],
                ["", "C"],
                ["A", "B"],
            ]
        );
        const { nodes } = visual.parseData(dv);
        // Only "A" and "B" should be created
        expect(nodes.size).toBe(2);
        expect(nodes.has("A")).toBe(true);
        expect(nodes.has("B")).toBe(true);
    });

    it("does not create a self-loop edge (eventId === nextStepId)", () => {
        const dv = makeTableDataView(
            [makeColumn("eventId"), makeColumn("nextStepId")],
            [["A", "A"]]
        );
        const { edges } = visual.parseData(dv);
        expect(edges.length).toBe(0);
    });
});

// ---------------------------------------------------------------------------
// computeLayout tests
// ---------------------------------------------------------------------------

describe("Visual.computeLayout", () => {
    let visual: Visual;

    beforeEach(() => {
        visual = makeVisual();
    });

    function buildNodes(ids: string[]): Map<string, ProcessNode> {
        const m = new Map<string, ProcessNode>();
        ids.forEach(id => m.set(id, { id, name: id, count: 1, layer: -1, posX: 0, posY: 0 }));
        return m;
    }

    it("assigns layer 0 to isolated nodes", () => {
        const nodes = buildNodes(["X", "Y"]);
        visual.computeLayout(nodes, [], 800, 600);
        expect(nodes.get("X")!.layer).toBe(0);
        expect(nodes.get("Y")!.layer).toBe(0);
    });

    it("assigns correct layers for a linear chain A→B→C", () => {
        const nodes = buildNodes(["A", "B", "C"]);
        const edges: ProcessEdge[] = [
            { sourceId: "A", targetId: "B", count: 1, avgDuration: null },
            { sourceId: "B", targetId: "C", count: 1, avgDuration: null },
        ];
        visual.computeLayout(nodes, edges, 800, 600);
        expect(nodes.get("A")!.layer).toBe(0);
        expect(nodes.get("B")!.layer).toBe(1);
        expect(nodes.get("C")!.layer).toBe(2);
    });

    it("places two roots (A and B) both on layer 0", () => {
        const nodes = buildNodes(["A", "B", "C"]);
        const edges: ProcessEdge[] = [
            { sourceId: "A", targetId: "C", count: 1, avgDuration: null },
            { sourceId: "B", targetId: "C", count: 1, avgDuration: null },
        ];
        visual.computeLayout(nodes, edges, 800, 600);
        expect(nodes.get("A")!.layer).toBe(0);
        expect(nodes.get("B")!.layer).toBe(0);
        expect(nodes.get("C")!.layer).toBe(1);
    });

    it("assigns positive posX and posY within viewport bounds", () => {
        const nodes = buildNodes(["A", "B"]);
        const edges: ProcessEdge[] = [
            { sourceId: "A", targetId: "B", count: 1, avgDuration: null },
        ];
        visual.computeLayout(nodes, edges, 800, 600);

        nodes.forEach(node => {
            expect(node.posX).toBeGreaterThan(0);
            expect(node.posX).toBeLessThanOrEqual(800);
            expect(node.posY).toBeGreaterThan(0);
            expect(node.posY).toBeLessThanOrEqual(600);
        });
    });

    it("handles a fully cyclic graph without throwing", () => {
        const nodes = buildNodes(["A", "B"]);
        const edges: ProcessEdge[] = [
            { sourceId: "A", targetId: "B", count: 1, avgDuration: null },
            { sourceId: "B", targetId: "A", count: 1, avgDuration: null },
        ];
        expect(() => visual.computeLayout(nodes, edges, 800, 600)).not.toThrow();
    });
});

// ---------------------------------------------------------------------------
// VisualSettings.parse tests
// ---------------------------------------------------------------------------

describe("VisualSettings.parse", () => {
    it("returns defaults when no metadata objects are present", () => {
        const dv = { metadata: { columns: [] } } as unknown as powerbi.DataView;
        const s = VisualSettings.parse(dv);
        expect(s.nodeSettings.nodeColor).toBe("#4472C4");
        expect(s.nodeSettings.nodeRadius).toBe(38);
        expect(s.nodeSettings.showCount).toBe(true);
        expect(s.edgeSettings.edgeColor).toBe("#888888");
        expect(s.labelSettings.fontSize).toBe(11);
    });

    it("reads nodeColor from metadata objects", () => {
        const dv = {
            metadata: {
                columns: [],
                objects: {
                    nodeSettings: {
                        nodeColor: { solid: { color: "#FF0000" } },
                    },
                },
            },
        } as unknown as powerbi.DataView;
        const s = VisualSettings.parse(dv);
        expect(s.nodeSettings.nodeColor).toBe("#FF0000");
    });

    it("clamps nodeRadius to minimum 10", () => {
        const dv = {
            metadata: {
                columns: [],
                objects: { nodeSettings: { nodeRadius: 2 } },
            },
        } as unknown as powerbi.DataView;
        const s = VisualSettings.parse(dv);
        expect(s.nodeSettings.nodeRadius).toBe(10);
    });

    it("clamps fontSize to minimum 8", () => {
        const dv = {
            metadata: {
                columns: [],
                objects: { labelSettings: { fontSize: 3 } },
            },
        } as unknown as powerbi.DataView;
        const s = VisualSettings.parse(dv);
        expect(s.labelSettings.fontSize).toBe(8);
    });
});
