# Power-BI-Visuals

Custom Visuals for Power BI

---

## Visuals

### 1. Process Mapping Tree (`processMappingTree/`)

Visualizes event-driven process flows as a directed graph. Bind an **Event
ID** column and an optional **Next Step ID** column to automatically build a
left-to-right flow diagram showing:

- Which events lead to which next steps.
- How many times each transition occurred (`× count` on edges).
- Average time between steps (when an **Avg Duration (ms)** measure is bound).

#### Data roles

| Field | Required | Type | Description |
|---|---|---|---|
| **Event ID** | ✅ | Text / Numeric | Unique identifier for each event type |
| **Next Step ID** | ➖ | Text / Numeric | ID of the following event (`null` = terminal step) |
| **Event Name** | ➖ | Text | Human-readable label shown inside each node |
| **Count** | ➖ | Numeric | Number of times this event or transition occurred |
| **Avg Duration (ms)** | ➖ | Numeric | Average milliseconds from this event to the next step |

#### Example data model

```
Events table
──────────────────────────────────────────────────────
event_id │ next_step_id │ event_name         │ count │ avg_duration_ms
─────────┼──────────────┼────────────────────┼───────┼────────────────
order    │ payment      │ Order Placed       │  1240 │  15000
payment  │ fulfillment  │ Payment Confirmed  │  1190 │  82000
fulfillment │ shipped   │ Fulfillment Start  │  1180 │ 180000
shipped  │ delivered    │ Shipped            │  1170 │ 604800000
delivered│ NULL         │ Delivered          │  1140 │ NULL
```

This produces a 5-node flow:
`Order Placed → Payment Confirmed → Fulfillment Start → Shipped → Delivered`

#### Formatting options

| Option | Default | Description |
|---|---|---|
| Node Color | `#4472C4` | Fill colour of each event node |
| Node Size (radius px) | `38` | Node circle radius |
| Show Event Count | `true` | Display `n = <count>` below each node |
| Edge Color | `#888888` | Colour of transition arrows |
| Show Avg Duration | `true` | Show `⌀ <time>` labels on edges |
| Font Size (px) | `11` | Label font size |
| Label Color | `#ffffff` | Label text colour |

#### Build & package

```bash
cd processMappingTree
npm install
npm run build         # produces processMappingTree.pbiviz
```

#### Run unit tests

```bash
cd processMappingTree
npm test
```

#### Import into Power BI

1. Open Power BI Desktop.
2. In the **Visualizations** pane click the three-dot menu (**…**) → **Import a visual from a file**.
3. Select the generated `processMappingTree.pbiviz` file.
4. The **Process Mapping Tree** icon will appear in the visualizations panel.
5. Drag it onto the report canvas and bind your data columns.
