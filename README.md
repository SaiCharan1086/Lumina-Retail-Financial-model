# Lumina Retail — Sales & Financial Analysis Model

This repository contains an end-to-end **Excel financial modeling project** for a simulated multi-region retail business (Lumina Retail Co.). The objective is to build a complete analytics workflow inside a single workbook: centralizing raw transaction data, rolling it up with formula-driven summaries, forecasting a 12-month income statement from an assumptions panel, and presenting results on an executive dashboard.

## Table of Contents

1. [Project Goals](#project-goals)
2. [Tech Stack](#tech-stack)
3. [Workbook Structure](#workbook-structure)
4. [End-to-End Model Pipeline](#end-to-end-model-pipeline)
   - [Raw Data Layer](#1-raw-data-layer)
   - [Lookup / Rate Card Layer](#2-lookup--rate-card-layer)
   - [Monthly Summary Layer](#3-monthly-summary-layer)
   - [Assumptions & Forecast Layer](#4-assumptions--forecast-layer)
   - [Dashboard Layer](#5-dashboard-layer)
5. [Executive Dashboard](#executive-dashboard)
6. [Key Results & Insights](#key-results--insights)
7. [How to Use This Model](#how-to-use-this-model)
8. [What This Project Demonstrates](#what-this-project-demonstrates)
9. [Roadmap (Future Enhancements)](#roadmap-future-enhancements)

---

## Project Goals

- Centralize two years of multi-region, multi-category retail transaction data in one workbook
- Build a `SUMIFS`-driven monthly rollup with margin % and YoY growth, with zero hardcoded outputs
- Drive a 12-month forward income statement entirely from a single, editable assumptions panel
- Compare three lookup approaches (`XLOOKUP`, `VLOOKUP`, `INDEX/MATCH`) solving the same problem
- Present results on a single-page executive dashboard with live KPI cards and charts

## Tech Stack

- **Excel** — modeling engine (formulas, no VBA/macros required)
- **SUMIFS / SUMIF / SUMPRODUCT** — data rollups and cross-tabulation
- **XLOOKUP / VLOOKUP / INDEX-MATCH** — category rate-card lookups, with `IFERROR` fallback for version compatibility
- **Native Excel charts** — dashboard visualizations (line, bar, pie)
- **Conditional formatting & color-coded conventions** — blue = input, black = formula, green = cross-sheet link

## Workbook Structure

```
Lumina Retail - Sales & Financial Model.xlsx
│
├── README                  # In-workbook guide to every sheet
├── Dashboard                # Executive KPI summary + 4 live charts — start here
├── Lookup Reference          # Category rate card; XLOOKUP vs VLOOKUP vs INDEX/MATCH
├── Sales Data                # 1,728-row raw transaction table (24 months x 4 regions x 3 categories)
├── Monthly Summary           # SUMIFS pivot-style rollup, margin %, YoY growth
├── Assumptions                # Editable driver panel (growth, margin, opex, tax)
└── Financial Model            # 12-month forward income statement
```

## End-to-End Model Pipeline

### 1. Raw Data Layer
The **Sales Data** sheet holds 1,728 rows of weekly-grain transactions across 24 months, 4 regions, and 3 categories. Unit Price is the only hardcoded (blue) input per row — every other column (Revenue, COGS, Gross Profit) is a live formula.

### 2. Lookup / Rate Card Layer
The **Lookup Reference** sheet stores a category rate card (base price, COGS rate, target margin) and demonstrates the same lookup solved three ways — `XLOOKUP` (wrapped in `IFERROR` with an `INDEX/MATCH` fallback for older Excel/LibreOffice), `VLOOKUP`, and `INDEX/MATCH` — with a live check confirming all three agree.

### 3. Monthly Summary Layer
`SUMIFS`/`SUMIF` roll the raw transactions up into a monthly table with Units Sold, Revenue, COGS, Gross Profit, Gross Margin %, YoY Growth, and a 3-month rolling average — plus a `SUMPRODUCT`-based Region × Category revenue breakdown.

### 4. Assumptions & Forecast Layer
A single **Assumptions** panel (growth rate, seasonality uplift, blended COGS %, opex %, D&A %, interest expense, tax rate) drives the entire **Financial Model** — a full 12-month income statement from Revenue down to Net Income, with zero hardcoded projections.

### 5. Dashboard Layer
The **Dashboard** sheet pulls live from Monthly Summary and Financial Model to surface six headline KPIs and four charts, all recalculating automatically when any input changes.

## Executive Dashboard

![Executive Dashboard](screenshots/dashboard.png)

## Key Results & Insights

- FY2025 revenue grew **16.0% YoY** with a blended gross margin of **53.1%**
- FY2026 is forecast to reach **₹48.8Cr** in revenue with a **17.2% net margin**
- Footwear carries the highest per-unit price but the tightest margin; Accessories carries the widest margin
- Nov/Dec seasonality uplift materially lifts Q4 forecast revenue and EBITDA

## How to Use This Model

1. Download `Lumina Retail - Sales & Financial Model.xlsx`
2. Open in Excel (2016+) or LibreOffice Calc
3. Start on the **Dashboard** sheet for the executive summary
4. To flex the forecast, edit any blue cell on the **Assumptions** sheet — the model, summary, and dashboard recalculate automatically
5. To explore the raw data, see the **Sales Data** sheet; to see how the rate card lookups work, see **Lookup Reference**

## What This Project Demonstrates

- Building a fully formula-driven financial model with zero hardcoded calculated outputs
- Designing a clean, auditable data pipeline (raw → lookup → summary → forecast → dashboard) inside Excel
- Practical fluency with `SUMIFS`, `SUMPRODUCT`, `XLOOKUP`, `VLOOKUP`, and `INDEX/MATCH`
- Driver-based financial forecasting (income statement modeling from a single assumptions panel)
- Dashboard design and data storytelling for an executive audience

## Roadmap (Future Enhancements)

- [ ] Add a Balance Sheet and Cash Flow Statement linked to the existing Income Statement
- [ ] Add scenario toggles (Base / Upside / Downside) via a dropdown-driven assumptions switch
- [ ] Add a Power Query-based refresh path for the raw data layer
- [ ] Publish an interactive Power BI / Tableau version of the dashboard

---
*Built as a demonstration of Excel-based financial modeling and data analysis skills.*
