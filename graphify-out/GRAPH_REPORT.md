# Graph Report - .  (2026-06-29)

## Corpus Check
- Corpus is ~26,782 words - fits in a single context window. You may not need a graph.

## Summary
- 161 nodes · 336 edges · 8 communities
- Extraction: 78% EXTRACTED · 22% INFERRED · 0% AMBIGUOUS · INFERRED: 75 edges (avg confidence: 0.58)
- Token cost: 143,333 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Admin Views & Purchasing|Admin Views & Purchasing]]
- [[_COMMUNITY_Camera P&L Model|Camera P&L Model]]
- [[_COMMUNITY_Rental Booking Logic|Rental Booking Logic]]
- [[_COMMUNITY_Sales & Variant Inventory|Sales & Variant Inventory]]
- [[_COMMUNITY_Frontend Templates & Concepts|Frontend Templates & Concepts]]
- [[_COMMUNITY_Flask Routes & Storefront|Flask Routes & Storefront]]
- [[_COMMUNITY_Database Seeding|Database Seeding]]

## God Nodes (most connected - your core abstractions)
1. `Camera` - 28 edges
2. `CamerasView` - 26 edges
3. `BookingGridView` - 16 edges
4. `StoreCostView` - 16 edges
5. `SheetView` - 16 edges
6. `CalendarView` - 14 edges
7. `FinanceView` - 13 edges
8. `parse_dt()` - 11 edges
9. `camera_pnl()` - 11 edges
10. `SaleRecord` - 11 edges

## Surprising Connections (you probably didn't know these)
- `Admin Shell (tintus_base.html)` --references--> `Flask + SQLAlchemy + Flask-Admin Stack`  [INFERRED]
  frontend/templates/admin/tintus_base.html → backend/requirements.txt
- `BookingGridView` --uses--> `Camera`  [INFERRED]
  backend/app.py → backend/models.py
- `BookingGridView` --uses--> `CameraVariant`  [INFERRED]
  backend/app.py → backend/models.py
- `BookingGridView` --uses--> `RentalBooking`  [INFERRED]
  backend/app.py → backend/models.py
- `BookingGridView` --uses--> `SaleRecord`  [INFERRED]
  backend/app.py → backend/models.py

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **Admin Dashboard Suite (all tintus_base children)** — frontend_templates_admin_booking_grid, frontend_templates_admin_calendar, frontend_templates_admin_cameras, frontend_templates_admin_chiphi, frontend_templates_admin_finance, frontend_templates_admin_sheet [EXTRACTED 1.00]
- **Customer Rental Booking Flow** — frontend_templates_store, frontend_templates_product_detail, frontend_templates_bang_gia, frontend_templates_rent [INFERRED 0.85]
- **Camera Sale / Purchase Flow** — frontend_templates_mua, frontend_templates_cart, frontend_templates_admin_cameras [INFERRED 0.75]

## Communities (8 total, 0 thin omitted)

### Community 0 - "Admin Views & Purchasing"
Cohesion: 0.13
Nodes (21): BookingGridView, CalendarView, checkout(), FinanceView, Shop overhead / operating costs (Chi phí tiệm) — rent, ads, shipping…, A free-form Google-Sheets-style scratch sheet., Google-Sheets-style grid: rows = cameras, columns = days, cells = bookings., FullCalendar week / 3-day time-grid view. (+13 more)

### Community 1 - "Camera P&L Model"
Cohesion: 0.08
Nodes (8): Camera, _fmt_vnd(), Capital tied up in unsold, still-sellable stock., Weighted-average unit cost across received import batches; falls back         to, Total income from all rental bookings of this camera., Per-unit sale: 1 if sold. Rental asset → binary is_sold., Sale: acquisition + repair. Rental: repairs only (gear is an owned         asset, Sale: realised only when sold; an unfixable unit is a write-off loss.

### Community 2 - "Rental Booking Logic"
Cohesion: 0.12
Nodes (15): bookings_overlapping(), color_map_for(), compute_rental_total(), parse_dt(), Map camera.id → hex color, stable by sort order., Parse a datetime from several accepted shapes; date-only → 09:00., Tiered total: day1 = price, day2 = price_d2, day3 = price_d3, day4+ = price_d4 e, Other bookings for the same camera that overlap [start, end). (+7 more)

### Community 3 - "Sales & Variant Inventory"
Cohesion: 0.17
Nodes (12): camera_pnl(), camera_stock(), CamerasView, Inventory manager. Renting tab = editable table; Selling tab = per-color/status, Mark a for-sale unit as sold, capturing price + customer info., serialize_receipt(), serialize_sale(), serialize_variant() (+4 more)

### Community 4 - "Frontend Templates & Concepts"
Cohesion: 0.18
Nodes (24): Backend Python Dependencies (Flask stack), Admin Dashboard (sidebar nav layout), Flask + SQLAlchemy + Flask-Admin Stack, Per-Camera Profit & Loss Model (revenue/cost/profit), Dual Inventory Model: Rent vs Sale Cameras, Rental Booking Flow, Camera Sale / Purchase Flow, Tiered Cumulative Daily Pricing (day1/d2/d3/d4+) (+16 more)

### Community 5 - "Flask Routes & Storefront"
Cohesion: 0.09
Nodes (9): add_to_cart(), buy_available(), buy_store(), cart(), ensure_schema(), Add any new columns / tables without dropping existing data., Public storefront for cameras that are for SALE (one inventory truth: variants)., A for-sale unit is buyable when in stock; rentals fall back to legacy stock. (+1 more)

### Community 6 - "Database Seeding"
Cohesion: 0.33
Nodes (5): backfill_costs(), Seed all camera products into the database. Run once: python seed.py Or called a, Insert any products not yet in the DB (idempotent)., Set import_cost on existing rows that don't have one yet (idempotent)., seed()

## Knowledge Gaps
- **3 isolated node(s):** `Backend Python Dependencies (Flask stack)`, `Admin Dashboard (sidebar nav layout)`, `Dual Inventory Model: Rent vs Sale Cameras`
  These have ≤1 connection - possible missing edges or undocumented components.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Camera` connect `Camera P&L Model` to `Admin Views & Purchasing`, `Rental Booking Logic`, `Sales & Variant Inventory`, `Database Seeding`?**
  _High betweenness centrality (0.252) - this node is a cross-community bridge._
- **Why does `CamerasView` connect `Sales & Variant Inventory` to `Admin Views & Purchasing`, `Camera P&L Model`, `Rental Booking Logic`, `Flask Routes & Storefront`?**
  _High betweenness centrality (0.130) - this node is a cross-community bridge._
- **Why does `SheetView` connect `Admin Views & Purchasing` to `Camera P&L Model`, `Rental Booking Logic`, `Sales & Variant Inventory`, `Flask Routes & Storefront`?**
  _High betweenness centrality (0.070) - this node is a cross-community bridge._
- **Are the 8 inferred relationships involving `Camera` (e.g. with `BookingGridView` and `CalendarView`) actually correct?**
  _`Camera` has 8 INFERRED edges - model-reasoned connections that need verification._
- **Are the 9 inferred relationships involving `CamerasView` (e.g. with `Camera` and `CameraVariant`) actually correct?**
  _`CamerasView` has 9 INFERRED edges - model-reasoned connections that need verification._
- **Are the 9 inferred relationships involving `BookingGridView` (e.g. with `Camera` and `CameraVariant`) actually correct?**
  _`BookingGridView` has 9 INFERRED edges - model-reasoned connections that need verification._
- **Are the 9 inferred relationships involving `StoreCostView` (e.g. with `Camera` and `CameraVariant`) actually correct?**
  _`StoreCostView` has 9 INFERRED edges - model-reasoned connections that need verification._