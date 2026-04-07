"use strict";

import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;

// ---------------------------------------------------------------------------
// Interfaces
// ---------------------------------------------------------------------------

export interface INodeSettings {
    nodeColor: string;
    nodeRadius: number;
    showCount: boolean;
}

export interface IEdgeSettings {
    edgeColor: string;
    showTimeMetrics: boolean;
}

export interface ILabelSettings {
    fontSize: number;
    fontColor: string;
}

// ---------------------------------------------------------------------------
// VisualSettings
// ---------------------------------------------------------------------------

export class VisualSettings {
    public nodeSettings: INodeSettings = {
        nodeColor: "#4472C4",
        nodeRadius: 38,
        showCount: true,
    };

    public edgeSettings: IEdgeSettings = {
        edgeColor: "#888888",
        showTimeMetrics: true,
    };

    public labelSettings: ILabelSettings = {
        fontSize: 11,
        fontColor: "#ffffff",
    };

    /**
     * Parse formatting settings from the Power BI DataView metadata objects.
     * Falls back to defaults for any missing or malformed values.
     */
    public static parse(dataView: DataView): VisualSettings {
        const settings = new VisualSettings();

        if (!dataView?.metadata?.objects) {
            return settings;
        }

        const objects = dataView.metadata.objects;

        // ── Node settings ──────────────────────────────────────────────────
        const ns = objects["nodeSettings"] as Record<string, powerbi.DataViewPropertyValue> | undefined;
        if (ns) {
            const fillColor = (ns["nodeColor"] as powerbi.Fill | undefined)?.solid?.color;
            if (fillColor) settings.nodeSettings.nodeColor = fillColor;

            const radius = ns["nodeRadius"];
            if (radius != null) settings.nodeSettings.nodeRadius = Math.max(10, Number(radius) || 38);

            const showCount = ns["showCount"];
            if (showCount != null) settings.nodeSettings.showCount = Boolean(showCount);
        }

        // ── Edge settings ──────────────────────────────────────────────────
        const es = objects["edgeSettings"] as Record<string, powerbi.DataViewPropertyValue> | undefined;
        if (es) {
            const fillColor = (es["edgeColor"] as powerbi.Fill | undefined)?.solid?.color;
            if (fillColor) settings.edgeSettings.edgeColor = fillColor;

            const showTime = es["showTimeMetrics"];
            if (showTime != null) settings.edgeSettings.showTimeMetrics = Boolean(showTime);
        }

        // ── Label settings ─────────────────────────────────────────────────
        const ls = objects["labelSettings"] as Record<string, powerbi.DataViewPropertyValue> | undefined;
        if (ls) {
            const size = ls["fontSize"];
            if (size != null) settings.labelSettings.fontSize = Math.max(8, Number(size) || 11);

            const fillColor = (ls["fontColor"] as powerbi.Fill | undefined)?.solid?.color;
            if (fillColor) settings.labelSettings.fontColor = fillColor;
        }

        return settings;
    }
}
