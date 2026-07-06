import json
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, timezone

db = SQLAlchemy()


class Camera(db.Model):
    id          = db.Column(db.Integer, primary_key=True)
    name        = db.Column(db.String(100), nullable=False)
    slug        = db.Column(db.String(100), unique=True)
    brand       = db.Column(db.String(50))
    type        = db.Column(db.String(20), default='Rent')   # 'Sale' or 'Rent'
    price       = db.Column(db.Integer, nullable=False)       # day-1 price (VND)
    price_d2    = db.Column(db.Integer, default=0)            # day-2 price (VND)
    price_d3    = db.Column(db.Integer, default=0)            # day-3 price (VND)
    price_d4    = db.Column(db.Integer, default=0)            # day 4+ price (VND)
    description = db.Column(db.Text)
    specs_json  = db.Column(db.Text)                          # JSON string
    stock       = db.Column(db.Integer, default=1)
    badge       = db.Column(db.String(50))
    featured    = db.Column(db.Boolean, default=False)
    category    = db.Column(db.String(40), default='')        # e.g. Mirrorless / Compact / Film (for Selling)

    # ── Inventory / financials (admin) ──────────────────────────────────────
    import_cost   = db.Column(db.Integer, default=0)          # default/fallback unit cost (VND)
    is_broken     = db.Column(db.Boolean, default=False)      # currently broken?
    repair_cost   = db.Column(db.Integer, default=0)          # cost to fix (VND)
    is_sold       = db.Column(db.Boolean, default=False)      # has been sold?
    sold_price    = db.Column(db.Integer, default=0)          # full agreed sale price (VND)
    deposit_amount = db.Column(db.Integer, default=0)         # cọc: cash received while state='deposit'
    sold_date     = db.Column(db.DateTime, nullable=True)
    reorder_point = db.Column(db.Integer, default=0)          # (legacy) low-stock threshold

    # ── Per-unit selling ledger (mirrors the shop's spreadsheet) ─────────────
    origin      = db.Column(db.String(20),  default='')       # 'Nhật' (N) / 'Việt' (V)
    accessory   = db.Column(db.String(120), default='')       # phụ kiện đi kèm (FLASH, pin…)
    color       = db.Column(db.String(40),  default='')       # màu máy
    date_in     = db.Column(db.DateTime, nullable=True)       # ngày máy về
    sale_state  = db.Column(db.String(20),  default='stock')  # stock | sold | fixing | unfixable
    sold_to     = db.Column(db.String(100), default='')       # người mua
    sold_phone  = db.Column(db.String(20),  default='')
    sold_source = db.Column(db.String(30),  default='')       # nguồn khách: Facebook/Instagram/Threads/TikTok…
    sold_note   = db.Column(db.String(300), default='')       # ghi chú khi bán (viết trong popup bán)
    gift_json   = db.Column(db.Text,        default='')       # JSON list of gifted accessories [{id,name,cost}]

    @property
    def specs(self):
        if self.specs_json:
            return json.loads(self.specs_json)
        return {}


    # ── Per-camera P&L (used by Finance + Cameras admin views) ───────────────
    @property
    def rental_revenue(self):
        """Total income from all rental bookings of this camera."""
        return int(sum((b.total_price or 0) for b in self.bookings))

    # A "deposit" (đã cọc) unit counts as realised revenue exactly like a sold one —
    # the customer has paid, the camera is reserved. It only lives in a separate tab.
    SOLD_STATES = ('sold', 'deposit')

    @property
    def units_sold(self):
        """Per-unit sale: 1 if sold/deposited. Rental asset → binary is_sold."""
        if self.type == 'Sale':
            return 1 if self.sale_state in self.SOLD_STATES else 0
        return 1 if self.is_sold else 0

    @property
    def sale_revenue(self):
        # A deposit (đã cọc) books the FULL sale price on the deposit day, same as a
        # completed sale. deposit_amount is tracked separately for reference only.
        if self.type == 'Sale':
            return int(self.sold_price or 0) if self.sale_state in self.SOLD_STATES else 0
        return int(self.sold_price or 0) if self.is_sold else 0

    @property
    def revenue(self):
        return self.rental_revenue + self.sale_revenue

    # ── Gifted accessories (bundled free with a camera sale) ─────────────────
    @property
    def gifts(self):
        if self.gift_json:
            try:
                return json.loads(self.gift_json)
            except (ValueError, TypeError):
                return []
        return []

    @property
    def gift_cost(self):
        """Total cost of accessories given away with this camera (eats profit)."""
        return int(sum(int(g.get('cost') or 0) for g in self.gifts))

    @property
    def gift_label(self):
        return ', '.join(g.get('name', '') for g in self.gifts if g.get('name'))

    @property
    def cost(self):
        """Sale: acquisition + repair + gifted accessories. Rental: repairs only
        (gear is an owned asset — its purchase price is not expensed).
        A deposit (đã cọc) is treated as a completed sale for accounting: full COGS
        is recognised on the deposit day, same as the full price (see sale_revenue)."""
        if self.type == 'Sale':
            return int((self.import_cost or 0) + (self.repair_cost or 0) + self.gift_cost)
        return int(self.repair_cost or 0)

    @property
    def profit(self):
        """Sale: realised only when sold/deposited; an unfixable unit is a write-off loss."""
        if self.type == 'Sale':
            if self.sale_state in self.SOLD_STATES:
                return self.sale_revenue - self.cost
            if self.sale_state == 'unfixable':
                return -self.cost
            return 0                       # stock / processing / fixing → not yet realised
        return self.revenue - self.cost

    @property
    def inventory_value(self):
        """Capital tied up in unsold, still-sellable stock. A deposit is booked as a
        completed sale (revenue on cọc day) so it is NOT counted as held inventory."""
        if self.type == 'Sale' and self.sale_state in ('stock', 'fixing', 'processing'):
            return int(self.import_cost or 0)
        return 0

    @property
    def is_broken_unit(self):
        return self.type == 'Sale' and self.sale_state in ('fixing', 'unfixable')

    @property
    def state_label(self):
        return {'stock': 'Còn hàng', 'processing': 'Đang xử lý', 'deposit': 'Đã cọc',
                'sold': 'Đã bán', 'fixing': 'Cần sửa',
                'unfixable': 'Không sửa được'}.get(self.sale_state or 'stock', 'Còn hàng')


    # Convenience: formatted price strings used in templates
    @property
    def price_d1_display(self):
        return _fmt_vnd(self.price)

    @property
    def price_d2_display(self):
        return _fmt_vnd(self.price_d2)

    @property
    def price_d3_display(self):
        return _fmt_vnd(self.price_d3)

    @property
    def price_d4_display(self):
        return f'+{_fmt_vnd(self.price_d4)}/ngày'


def _fmt_vnd(value):
    if not value:
        return '—'
    v = int(value)
    if v >= 1000:
        return f'{v // 1000}K'
    return str(v)


class RentalBooking(db.Model):
    id            = db.Column(db.Integer, primary_key=True)
    camera_id     = db.Column(db.Integer, db.ForeignKey('camera.id'))
    customer_name = db.Column(db.String(100))
    phone         = db.Column(db.String(20))
    notes         = db.Column(db.Text)
    start_date    = db.Column(db.DateTime, nullable=False)
    end_date      = db.Column(db.DateTime, nullable=False)
    total_price   = db.Column(db.Float)
    camera        = db.relationship('Camera', backref='bookings')


class Appointment(db.Model):
    """A scheduled in-shop appointment for a client to come buy a camera."""
    id            = db.Column(db.Integer, primary_key=True)
    customer_name = db.Column(db.String(100))
    phone         = db.Column(db.String(20))
    start_time    = db.Column(db.DateTime, nullable=False)      # hẹn gặp lúc
    end_time      = db.Column(db.DateTime)                      # optional
    camera_id     = db.Column(db.Integer, db.ForeignKey('camera.id'), nullable=True)  # optional target unit
    interest      = db.Column(db.String(200), default='')       # máy / nhu cầu muốn mua
    notes         = db.Column(db.Text)
    status        = db.Column(db.String(20), default='booked')  # booked | done | cancelled
    created_at    = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    camera        = db.relationship('Camera')


class Lead(db.Model):
    """A rental/purchase inquiry captured from the public site (homepage form).
    Not yet a booking — a to-follow-up contact so no walk-in inquiry is lost."""
    id            = db.Column(db.Integer, primary_key=True)
    customer_name = db.Column(db.String(100), default='')
    phone         = db.Column(db.String(30),  default='')
    notes         = db.Column(db.Text,        default='')
    start_date    = db.Column(db.DateTime, nullable=True)     # desired rental start
    end_date      = db.Column(db.DateTime, nullable=True)     # desired rental end
    source        = db.Column(db.String(40),  default='Trang chủ')
    status        = db.Column(db.String(20),  default='new')  # new | contacted | done
    created_at    = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))


class RepairLog(db.Model):
    """A dated rental-camera repair expense, logged whenever a rental unit's cumulative
    repair_cost is edited, so the cost lands in the Finance period it was incurred.
    (Sale-camera repairs are already expensed in COGS at sale time, so they're not logged.)
    created_at is naive local time to match Finance's naive date-window comparisons."""
    id         = db.Column(db.Integer, primary_key=True)
    camera_id  = db.Column(db.Integer, db.ForeignKey('camera.id'), nullable=False)
    amount     = db.Column(db.Integer, default=0)             # delta added this edit (may be negative)
    note       = db.Column(db.String(200), default='')
    created_at = db.Column(db.DateTime, default=datetime.now)


class ActivityLog(db.Model):
    """One user-visible admin action, with enough row-level before/after data to
    reverse (undo) and re-apply (redo) it. Row changes are captured automatically
    via SQLAlchemy session events and grouped per request into a single entry.

    `changes_json` is a JSON list of {action, table, pk, before, after}; datetimes
    are wrapped as {"__dt__": iso} so they round-trip losslessly. `undone` marks
    whether this entry is currently rolled back — the applied entries always form a
    contiguous prefix (by id) and the undone ones a contiguous suffix, giving a
    linear undo/redo timeline."""
    id           = db.Column(db.Integer, primary_key=True)
    label        = db.Column(db.String(300), default='')   # human summary, Vietnamese
    changes_json = db.Column(db.Text, default='')
    undone       = db.Column(db.Boolean, default=False)
    created_at   = db.Column(db.DateTime, default=datetime.now)   # naive local time


class StoreCost(db.Model):
    """A shop overhead / operating cost (rent, ads, shipping…) — 'Chi phí tiệm'."""
    id         = db.Column(db.Integer, primary_key=True)
    category   = db.Column(db.String(60), default='Khác')     # LOẠI: Tiền nhà / Quảng cáo / Vận chuyển…
    note       = db.Column(db.String(200), default='')
    amount     = db.Column(db.Integer, default=0)             # VND
    cost_date  = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))


class Accessory(db.Model):
    """A non-camera item sold by the quantity — lens, pin, thẻ nhớ, sạc…"""
    id         = db.Column(db.Integer, primary_key=True)
    name       = db.Column(db.String(120), nullable=False)
    category   = db.Column(db.String(40),  default='')        # Lens / Pin / Thẻ nhớ…
    cost       = db.Column(db.Integer, default=0)             # giá nhập (also the gift deduction)
    price      = db.Column(db.Integer, default=0)             # giá bán mặc định
    stock      = db.Column(db.Integer, default=0)             # số lượng tồn kho
    note       = db.Column(db.String(200), default='')
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    @property
    def stock_value(self):
        return int((self.stock or 0) * (self.cost or 0))


class AccessorySale(db.Model):
    """A logged sale of N units of an accessory (feeds Finance)."""
    id           = db.Column(db.Integer, primary_key=True)
    accessory_id = db.Column(db.Integer, db.ForeignKey('accessory.id'), nullable=True)
    name         = db.Column(db.String(120), default='')      # snapshot of the name at sale time
    quantity     = db.Column(db.Integer, default=1)
    unit_price   = db.Column(db.Integer, default=0)
    unit_cost    = db.Column(db.Integer, default=0)           # cost snapshot (COGS)
    note         = db.Column(db.String(200), default='')      # optional note per sale
    customer_name = db.Column(db.String(100), default='')     # tên khách
    phone         = db.Column(db.String(20),  default='')     # sđt khách
    sale_date    = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    @property
    def revenue(self):
        return int((self.quantity or 0) * (self.unit_price or 0))

    @property
    def cost_total(self):
        return int((self.quantity or 0) * (self.unit_cost or 0))

    @property
    def profit(self):
        return self.revenue - self.cost_total


class PurchaseOrder(db.Model):
    id            = db.Column(db.Integer, primary_key=True)
    camera_id     = db.Column(db.Integer, db.ForeignKey('camera.id'))
    customer_name = db.Column(db.String(100), nullable=False)
    phone         = db.Column(db.String(20),  nullable=False)
    address       = db.Column(db.Text,        nullable=False)
    quantity      = db.Column(db.Integer,     nullable=False, default=1)
    total_price   = db.Column(db.Float,       nullable=False)
    order_date    = db.Column(db.DateTime,    default=lambda: datetime.now(timezone.utc))


class Sheet(db.Model):
    """A free-form spreadsheet (Google-Sheets-style) stored as a JSON 2D array."""
    id         = db.Column(db.Integer, primary_key=True)
    name       = db.Column(db.String(100), default='Sổ tay')
    data_json  = db.Column(db.Text)   # JSON: list of rows, each a list of cell strings
    updated_at = db.Column(db.DateTime,
                           default=lambda: datetime.now(timezone.utc),
                           onupdate=lambda: datetime.now(timezone.utc))

    @property
    def data(self):
        if self.data_json:
            try:
                return json.loads(self.data_json)
            except (ValueError, TypeError):
                return []
        return []
