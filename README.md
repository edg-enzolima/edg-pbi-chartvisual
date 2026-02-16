# EDG Chart Visual (Power BI Custom Visual)

This repository contains a Power BI custom visual project at:

- `diagonalStripeBarChart`

## What this visual does

- Renders a bar chart on an HTML canvas.
- Supports diagonal stripes fixed at **45 degrees** (not configurable).
- Supports two stripe modes in the format pane:
  - **All bars**: every bar gets diagonal stripes.
  - **Forecast**: stripes apply only when the X axis is a date field, and only for **today/future** dates.
    - Past dates remain solid.

## Local development

```bash
cd diagonalStripeBarChart
npm install
npm run start
```

## Build the `.pbiviz` package

```bash
cd diagonalStripeBarChart
npm run package
```
