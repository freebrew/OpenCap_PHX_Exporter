---
name: canvas
description: >-
  Render standalone analytical output as a live React app beside the chat.
  Use when producing field cap summaries, slide sheet metrics, daily paperwork
  reports, cost tables, import results, well-log comparisons, or any data-heavy
  deliverable for the PHX_FieldCap project. If you catch yourself writing a
  markdown table with more than a handful of rows, use a canvas instead.
---

# Canvas — PHX_FieldCap

## When to use

- Field cap daily/weekly summary tables
- Slide sheet import results (row counts, error lists, column mappings)
- Well cost / AFE comparison views
- MDL_Setup module audit outputs (procedure list, line counts, etc.)
- Any report that benefits from visual hierarchy over a flat markdown dump

## When NOT to use

- Quick one-line answers
- Code edits or VBA diffs
- Active debugging output mid-session

## Location rule

Canvases are written to:
```
~/.cursor/projects/<workspace>/canvases/<name>.canvas.tsx
```
Never create subfolders or helper files. One `.canvas.tsx` per canvas.

## Import rule

```tsx
import { Divider, Grid, H1, H2, Stack, Stat, Table, Text } from 'cursor/canvas';
```
Only import from `cursor/canvas`. No fetch calls, no npm packages, all data inline.

## Design constraints

- No gradients, no emojis, no box-shadows, no rainbow colors
- Use `useHostTheme()` tokens — no hardcoded hex
- Stats in a `Grid`, data in `Table`, text sections without wrapping every card

## Minimal field-cap pattern

```tsx
import { Divider, Grid, H1, H2, Stack, Stat, Table, Text } from 'cursor/canvas';

export default function FieldCapSummary() {
  return (
    <Stack gap={20}>
      <H1>Daily Field Cap — TOURMALINE HZ SUNDOWN H04-04-077-17 P3</H1>
      <Grid columns={3} gap={16}>
        <Stat value="142" label="Rows Imported" />
        <Stat value="3" label="Errors" tone="warning" />
        <Stat value="139" label="Processed" tone="success" />
      </Grid>
      <Divider />
      <H2>Import Log</H2>
      <Table
        headers={["Row", "Field", "Value", "Status"]}
        rows={[
          ["1", "WellID", "H04-04-077-17", "OK"],
          ["2", "Phase", "P3", "OK"],
          ["3", "Cost", "", "Missing"],
        ]}
        rowTone={[undefined, undefined, "warning"]}
      />
      <Text tone="secondary" size="small">Generated: May 11, 2026</Text>
    </Stack>
  );
}
```
