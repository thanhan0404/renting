import os
import io
import math
import json
from datetime import datetime, timedelta, timezone
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
from flask_admin import Admin, BaseView, expose, AdminIndexView
from flask_admin.contrib.sqla import ModelView
from sqlalchemy import inspect as sa_inspect, text
from werkzeug.security import generate_password_hash, check_password_hash

from models import (db, Camera, RentalBooking, PurchaseOrder, Sheet,
                    CameraVariant, SaleRecord, Supplier, StockReceipt, StoreCost,
                    Accessory, AccessorySale, Appointment)
from seed import seed as seed_db, backfill_costs

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:                       # pragma: no cover
    HAS_OPENPYXL = False

# Camera categories for the Selling tab "type" dropdown.
CAMERA_CATEGORIES = ['Mirrorless', 'Compact', 'DSLR', 'Máy phim', 'Action Cam',
                     'Flycam', 'Chụp lấy liền', 'Khác']

# Where a buying customer came from (sale popup dropdown).
SALE_SOURCES = ['Facebook', 'Instagram', 'Threads', 'TikTok', 'Tại shop', 'Khác']

# Accessory / other-goods categories (Phụ kiện tab).
ACCESSORY_CATEGORIES = ['Lens', 'Pin', 'Thẻ nhớ', 'Sạc', 'Túi/Bao', 'Dây đeo',
                        'Tripod', 'Filter', 'Khác']

# Column order of the Excel import template (mirrors the shop's spreadsheet).
IMPORT_COLUMNS = ['Tên máy', 'Màu', 'Dòng máy', 'Nơi (N/V)', 'Phụ kiện',
                  'Ngày về (YYYY-MM-DD)', 'Ghi chú', 'Giá nhập', 'Giá bán', 'Trạng thái']

# state text (Vietnamese / English) → internal sale_state
STATE_ALIASES = {
    'còn hàng': 'stock', 'con hang': 'stock', 'stock': 'stock', 'tồn': 'stock', '': 'stock',
    'đã bán': 'sold', 'da ban': 'sold', 'sold': 'sold', 'bán': 'sold',
    'cần sửa': 'fixing', 'can sua': 'fixing', 'fixing': 'fixing', 'sửa': 'fixing', 'hỏng': 'fixing',
    'không sửa được': 'unfixable', 'khong sua duoc': 'unfixable', 'unfixable': 'unfixable',
}


def normalize_origin(value):
    """Map a free-text origin cell to 'Nhật' / 'Việt' / ''."""
    v = (str(value or '')).strip().lower()
    if v in ('n', 'nhật', 'nhat', 'japan', 'jp', 'nhật bản', 'nhat ban'):
        return 'Nhật'
    if v in ('v', 'việt', 'viet', 'vietnam', 'vn', 'việt nam', 'viet nam'):
        return 'Việt'
    return ''


def normalize_state(value):
    return STATE_ALIASES.get((str(value or '')).strip().lower(), 'stock')


def _to_int(value):
    """Coerce a spreadsheet cell (number / '1.200.000' / '1,200,000') to int."""
    if value is None or value == '':
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    digits = ''.join(ch for ch in str(value) if ch.isdigit())
    return int(digits) if digits else 0


def _cell_date(value):
    if isinstance(value, datetime):
        return value
    return parse_dt(str(value)) if value else None

basedir = os.path.abspath(os.path.dirname(__file__))
# Templates and static assets live in the sibling frontend/ folder.
frontend_dir = os.path.join(basedir, '..', 'frontend')
template_dir = os.path.join(frontend_dir, 'templates')
static_dir   = os.path.join(frontend_dir, 'static')

app = Flask(__name__, template_folder=template_dir,
            static_folder=static_dir, static_url_path='/static')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///camerashop.db'
# Session-signing key. MUST be set via env in production (a leaked/default key
# lets anyone forge an admin session cookie).
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-insecure-change-me')
app.config['TEMPLATES_AUTO_RELOAD'] = True   # pick up template edits without a restart
app.jinja_env.auto_reload = True

# ── Admin security ───────────────────────────────────────────────────────────
# The whole admin panel lives behind a secret, hard-to-guess URL prefix AND a
# login. Both are configurable via environment variables (see SECURITY.md).
ADMIN_URL_PREFIX = '/' + os.environ.get('ADMIN_URL_PREFIX', 'quanly-tintus').strip('/')
ADMIN_USERNAME   = os.environ.get('ADMIN_USERNAME', 'admin')
ADMIN_PASSWORD_HASH = generate_password_hash(os.environ.get('ADMIN_PASSWORD', 'tintus@2026'))
LOGIN_PATH  = ADMIN_URL_PREFIX + '/login'
LOGOUT_PATH = ADMIN_URL_PREFIX + '/logout'

# Harden the session cookie. Set SESSION_COOKIE_SECURE=1 once you only serve
# the admin over HTTPS (e.g. behind Cloudflare) — leave 0 for plain-HTTP LAN access.
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=(os.environ.get('SESSION_COOKIE_SECURE', '0') == '1'),
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
)

db.init_app(app)


class RedirectIndexView(AdminIndexView):
    """The default Flask-Admin home is unused → send it straight to Cameras."""
    @expose('/')
    def index(self):
        return redirect(url_for('cameras.index'))


admin = Admin(app, name='tintus.digicam Admin',
              index_view=RedirectIndexView(url=ADMIN_URL_PREFIX, endpoint='admin'))
admin.add_view(ModelView(Camera, db.session))
admin.add_view(ModelView(RentalBooking, db.session))
admin.add_view(ModelView(PurchaseOrder, db.session))


# ── Gate every admin request behind the login ────────────────────────────────
@app.before_request
def _require_admin_login():
    p = request.path
    if not (p == ADMIN_URL_PREFIX or p.startswith(ADMIN_URL_PREFIX + '/')):
        return                                  # public site → no gate
    if p in (LOGIN_PATH, LOGOUT_PATH):
        return                                  # auth endpoints stay reachable
    if not session.get('is_admin'):
        return redirect(url_for('admin_login', next=p))


@app.route(LOGIN_PATH, methods=['GET', 'POST'])
def admin_login():
    if session.get('is_admin'):
        return redirect(url_for('cameras.index'))
    error = None
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        if username == ADMIN_USERNAME and check_password_hash(ADMIN_PASSWORD_HASH, password):
            session.clear()
            session['is_admin'] = True
            session['admin_user'] = username
            session.permanent = True
            nxt = request.args.get('next') or url_for('cameras.index')
            if not nxt.startswith(ADMIN_URL_PREFIX):     # block open-redirects
                nxt = url_for('cameras.index')
            return redirect(nxt)
        error = 'Sai tài khoản hoặc mật khẩu.'
    return render_template('admin_login.html', error=error, login_action=LOGIN_PATH)


@app.route(LOGOUT_PATH)
def admin_logout():
    session.clear()
    return redirect(LOGIN_PATH)


# ── Jinja2 filter: 280000 → "280K" ─────────────────────────────────────────
@app.template_filter('vnd')
def vnd_filter(value):
    """Full VND with Vietnamese thousands separators: 280000 → '280.000' (— if empty)."""
    if not value:
        return '—'
    return '{:,.0f}'.format(int(value)).replace(',', '.')


# ── Jinja2 filter: 1200000 → "1.200.000" (Vietnamese thousands separators) ──
@app.template_filter('money')
def money_filter(value):
    return '{:,.0f}'.format(int(value or 0)).replace(',', '.')


# ── Global template context: cart badge count on every page ─────────────────
@app.context_processor
def inject_cart_count():
    return {'cart_count': sum(session.get('cart', {}).values())}


# ── Brand display config (used in templates) ────────────────────────────────
BRAND_CONFIG = {
    'Canon':   {'header': 'bg-red-700',   'hover': 'hover:bg-red-50',   'text': 'text-red-700',   'icon': 'bg-red-100',   'emoji': '📷'},
    'Fujifilm':{'header': 'bg-green-700', 'hover': 'hover:bg-green-50', 'text': 'text-green-700', 'icon': 'bg-green-100', 'emoji': '📷'},
    'Sony':    {'header': 'bg-blue-700',  'hover': 'hover:bg-blue-50',  'text': 'text-blue-700',  'icon': 'bg-blue-100',  'emoji': '📷'},
    'DJI':     {'header': 'bg-gray-700',  'hover': 'hover:bg-gray-50',  'text': 'text-gray-700',  'icon': 'bg-gray-100',  'emoji': '🎥'},
    'Lumix':   {'header': 'bg-indigo-700','hover': 'hover:bg-indigo-50','text': 'text-indigo-700','icon': 'bg-indigo-100','emoji': '📷'},
    'Casio':   {'header': 'bg-yellow-600','hover': 'hover:bg-yellow-50','text': 'text-yellow-700','icon': 'bg-yellow-100','emoji': '📷'},
    'Nikon':   {'header': 'bg-yellow-700','hover': 'hover:bg-yellow-50','text': 'text-yellow-700','icon': 'bg-yellow-100','emoji': '📷'},
}


# ── ROUTES ──────────────────────────────────────────────────────────────────

@app.route('/')
def home():
    featured = Camera.query.filter_by(type='Rent', featured=True).order_by(Camera.price.desc()).limit(8).all()
    return render_template('index.html', featured=featured)


@app.route('/store')
def store():
    brand = request.args.get('brand', '').strip()
    keyword = request.args.get('q', '').strip()
    price = request.args.get('price', '').strip()
    sort = request.args.get('sort', '').strip()

    query = Camera.query.filter_by(type='Rent')
    if brand:
        query = query.filter(Camera.brand.ilike(brand))
    if keyword:
        like = f'%{keyword}%'
        query = query.filter(db.or_(Camera.name.ilike(like), Camera.brand.ilike(like)))
    if price == 'low':
        query = query.filter(Camera.price < 200000)
    elif price == 'mid':
        query = query.filter(Camera.price >= 200000, Camera.price <= 350000)
    elif price == 'high':
        query = query.filter(Camera.price > 350000)

    if sort == 'price_asc':
        query = query.order_by(Camera.price.asc())
    elif sort == 'price_desc':
        query = query.order_by(Camera.price.desc())
    elif sort == 'name':
        query = query.order_by(Camera.name.asc())
    else:
        query = query.order_by(Camera.brand, Camera.price.desc())

    cameras = query.all()
    brands = [r[0] for r in db.session.query(Camera.brand).filter_by(type='Rent').distinct().order_by(Camera.brand).all()]
    return render_template('store.html', cameras=cameras, current_brand=brand, brands=brands,
                           keyword=keyword, current_price=price, current_sort=sort)


@app.route('/mua')
def buy_store():
    """Public storefront for cameras that are for SALE (one inventory truth: variants)."""
    brand = request.args.get('brand', '').strip()
    keyword = request.args.get('q', '').strip()
    query = Camera.query.filter_by(type='Sale', sale_state='stock')   # only available units
    if brand:
        query = query.filter(Camera.brand.ilike(brand))
    if keyword:
        like = f'%{keyword}%'
        query = query.filter(db.or_(Camera.name.ilike(like), Camera.brand.ilike(like)))
    cameras = query.order_by(Camera.brand, Camera.price.desc()).all()
    brands = [r[0] for r in db.session.query(Camera.brand).filter_by(type='Sale', sale_state='stock').distinct().order_by(Camera.brand).all() if r[0]]
    return render_template('mua.html', cameras=cameras, current_brand=brand, brands=brands, keyword=keyword)


@app.route('/san-pham')
def san_pham():
    return redirect(url_for('store'))


@app.route('/bang-gia')
def bang_gia():
    cameras = Camera.query.filter_by(type='Rent').order_by(Camera.brand, Camera.price.desc()).all()
    # Group by brand preserving order
    grouped = {}
    for cam in cameras:
        grouped.setdefault(cam.brand, []).append(cam)
    return render_template('bang_gia.html', grouped=grouped, brand_config=BRAND_CONFIG)


@app.route('/bang-gia-cho-thue-may-anh-2saigon')
def bang_gia_alias():
    return redirect(url_for('bang_gia'))


@app.route('/product/<slug>')
def product_detail(slug):
    product = Camera.query.filter_by(slug=slug).first_or_404()
    related = (Camera.query
               .filter_by(brand=product.brand, type='Rent')
               .filter(Camera.id != product.id)
               .limit(4).all())
    if len(related) < 4:
        others = (Camera.query
                  .filter(Camera.brand != product.brand, Camera.type == 'Rent')
                  .limit(4 - len(related)).all())
        related += others
    return render_template('product_detail.html', product=product, related=related)


@app.route('/<slug>')
def product_alias(slug):
    cam = Camera.query.filter_by(slug=slug).first()
    if cam:
        return redirect(url_for('product_detail', slug=slug))
    return redirect(url_for('home'))


# ── Cart (online PURCHASE of for-sale cameras → SaleRecord ledger) ────────────

def buy_available(camera):
    """A for-sale unit is buyable when in stock; rentals fall back to legacy stock."""
    if camera.type == 'Sale':
        return 1 if camera.sale_state == 'stock' else 0
    return camera.stock or 0


@app.route('/add_to_cart/<int:camera_id>')
def add_to_cart(camera_id):
    if 'cart' not in session:
        session['cart'] = {}
    cart = session['cart']
    camera = Camera.query.get_or_404(camera_id)
    avail = buy_available(camera)
    if avail > 0:
        key = str(camera_id)
        if cart.get(key, 0) < avail:
            cart[key] = cart.get(key, 0) + 1
            session.modified = True
    return redirect(url_for('cart'))


@app.route('/cart')
def cart():
    items, total = [], 0
    for cam_id, qty in session.get('cart', {}).items():
        camera = db.session.get(Camera, int(cam_id))
        if camera:
            item_total = camera.price * qty
            total += item_total
            items.append({'camera': camera, 'quantity': qty, 'item_total': item_total,
                          'available': buy_available(camera)})
    return render_template('cart.html', items=items, total=total)


@app.route('/checkout', methods=['POST'])
def checkout():
    cart_items = session.get('cart', {})
    if not cart_items:
        return redirect(url_for('buy_store'))
    customer_name = request.form.get('customer_name')
    phone         = request.form.get('phone')
    address       = request.form.get('address')
    for cam_id, qty in cart_items.items():
        camera = db.session.get(Camera, int(cam_id))
        if not camera:
            continue
        if camera.type == 'Sale':
            if camera.sale_state == 'stock':       # one physical unit → mark sold
                camera.sale_state = 'sold'
                camera.is_sold = True
                camera.sold_price = camera.price
                camera.sold_to = customer_name
                camera.sold_phone = phone
                camera.sold_date = datetime.now()
        elif (camera.stock or 0) >= qty:           # legacy: buying a rental asset
            camera.stock -= qty
            db.session.add(PurchaseOrder(
                camera_id=camera.id, customer_name=customer_name, phone=phone,
                address=address, quantity=qty, total_price=camera.price * qty))
    db.session.commit()
    session.pop('cart', None)
    return render_template('cart.html',
                           message='Mua hàng thành công! Chúng tôi sẽ sớm liên hệ với bạn.',
                           items=[], total=0)


# ── Rental API ───────────────────────────────────────────────────────────────

@app.route('/api/camera-availability/<int:camera_id>')
def camera_availability_api(camera_id):
    """Public: only reveals whether a camera is free for a given [start,end) — not the schedule."""
    start = parse_dt(request.args.get('start'))
    end   = parse_dt(request.args.get('end'))
    if not start or not end or end <= start:
        return jsonify({'available': False, 'error': 'invalid range'}), 400
    busy = bool(bookings_overlapping(start, end, camera_id)) or \
           bool(appointments_overlapping(start, end, camera_id))
    return jsonify({'available': not busy})


@app.route('/api/available-cameras')
def available_cameras_api():
    start_str = request.args.get('start')
    end_str   = request.args.get('end')
    start = parse_dt(start_str)
    end   = parse_dt(end_str)
    if not start or not end:
        return jsonify({'error': 'Invalid date format'}), 400
    if end <= start:
        return jsonify({'error': 'End date must be after start date'}), 400
    busy_ids = [r[0] for r in db.session.query(RentalBooking.camera_id).filter(
        RentalBooking.start_date < end,
        RentalBooking.end_date > start,
    ).distinct().all()]
    q = Camera.query.filter_by(type='Rent')
    if busy_ids:
        q = q.filter(~Camera.id.in_(busy_ids))
    return jsonify([{'id': c.id, 'name': c.name, 'price': c.price} for c in q.all()])


@app.route('/rent', methods=['GET', 'POST'])
def rent():
    message = error = None
    if request.method == 'POST':
        camera = db.session.get(Camera, request.form.get('camera_id'))
        start_date = parse_dt(request.form.get('start_date'))
        end_date   = parse_dt(request.form.get('end_date'))
        if not camera or camera.type != 'Rent':
            error = 'Vui lòng chọn một máy hợp lệ để thuê.'
        elif not start_date or not end_date or end_date <= start_date:
            error = 'Thời gian thuê không hợp lệ — giờ trả phải sau giờ nhận.'
        elif bookings_overlapping(start_date, end_date, camera.id) or appointments_overlapping(start_date, end_date, camera.id):
            error = f'Máy “{camera.name}” đã có người đặt trong khung giờ bạn chọn. Vui lòng chọn thời gian khác hoặc máy khác.'
        else:
            total_price = compute_rental_total(camera, start_date, end_date)
            b = RentalBooking(
                camera_id=camera.id, customer_name=(request.form.get('customer_name') or '').strip(),
                phone=(request.form.get('phone') or '').strip(), notes=(request.form.get('notes') or '').strip(),
                start_date=start_date, end_date=end_date, total_price=total_price,
            )
            db.session.add(b)
            db.session.commit()
            race_err = confirm_rental_or_rollback(b)
            if race_err:
                error = race_err
            else:
                message = f'Thành công! Bạn đã đặt thuê {camera.name}. Tổng tiền dự kiến: {total_price:,.0f}đ'.replace(',', '.')
    rent_cameras = rent_cameras_ordered()
    camera_slug = request.args.get('camera', '')
    selected_camera = Camera.query.filter_by(slug=camera_slug).first() if camera_slug else None
    return render_template('rent.html', cameras=rent_cameras, message=message,
                           error=error, selected_camera=selected_camera)


# ── Admin helpers ─────────────────────────────────────────────────────────────

# Stable palette: each camera keeps the same color everywhere in the admin.
ADMIN_COLORS = [
    '#ef4444', '#3b82f6', '#10b981', '#f59e0b', '#8b5cf6',
    '#f97316', '#06b6d4', '#ec4899', '#84cc16', '#6366f1',
    '#14b8a6', '#e11d48', '#0ea5e9', '#a855f7', '#22c55e',
]

APPT_COLOR = '#0e7490'   # buy-appointments: a single high-contrast teal (vs per-camera rental colors)


def color_map_for(cameras):
    """Map camera.id → hex color, stable by sort order."""
    return {cam.id: ADMIN_COLORS[i % len(ADMIN_COLORS)] for i, cam in enumerate(cameras)}


def rent_cameras_ordered(include_sold=False):
    q = Camera.query.filter_by(type='Rent')
    if not include_sold:
        q = q.filter(Camera.is_sold.isnot(True))
    return q.order_by(Camera.brand, Camera.name).all()


def parse_dt(value):
    """Parse a datetime from several accepted shapes; date-only → 09:00."""
    if not value:
        return None
    value = value.strip()
    for fmt in ('%Y-%m-%dT%H:%M', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    try:
        d = datetime.strptime(value, '%Y-%m-%d')
        return d.replace(hour=9, minute=0)
    except ValueError:
        pass
    try:
        return datetime.fromisoformat(value[:19])
    except ValueError:
        return None


def compute_rental_total(camera, start, end):
    """Tiered total: day1 = price, day2 = price_d2, day3 = price_d3, day4+ = price_d4 each."""
    if not camera or not start or not end:
        return 0
    seconds = (end - start).total_seconds()
    days = max(1, math.ceil(seconds / 86400)) if seconds > 0 else 1
    total = camera.price or 0
    if days >= 2:
        total += camera.price_d2 or 0
    if days >= 3:
        total += camera.price_d3 or 0
    if days > 3:
        total += (camera.price_d4 or 0) * (days - 3)
    return int(total)


def serialize_booking(b, color=None):
    cam = b.camera
    days = max(1, math.ceil((b.end_date - b.start_date).total_seconds() / 86400))
    return {
        'id':         b.id,
        'camera_id':  b.camera_id,
        'camera':     cam.name if cam else '—',
        'brand':      cam.brand if cam else '',
        'customer':   b.customer_name or '',
        'phone':      b.phone or '',
        'notes':      b.notes or '',
        'total':      int(b.total_price or 0),
        'days':       days,
        'start':      b.start_date.isoformat(),
        'end':        b.end_date.isoformat(),
        'start_date': b.start_date.strftime('%Y-%m-%d'),
        'end_date':   b.end_date.strftime('%Y-%m-%d'),
        'color':      color or '#6b7280',
    }


def serialize_appt(a):
    cam = a.camera
    end = a.end_time or (a.start_time + timedelta(minutes=45))
    return {
        'id':        a.id,
        'customer':  a.customer_name or '',
        'phone':     a.phone or '',
        'camera_id': a.camera_id,
        'camera':    cam.name if cam else '',
        'interest':  a.interest or '',
        'notes':     a.notes or '',
        'status':    a.status or 'booked',
        'start':     a.start_time.isoformat(),
        'end':       end.isoformat(),
        'date':      a.start_time.strftime('%Y-%m-%d'),
        'time':      a.start_time.strftime('%H:%M'),
        'kind':      'appt',
    }


def bookings_overlapping(start, end, camera_id, exclude_id=None):
    """Other bookings for the same camera that overlap [start, end)."""
    q = RentalBooking.query.filter(
        RentalBooking.camera_id == camera_id,
        RentalBooking.start_date < end,
        RentalBooking.end_date > start,
    )
    if exclude_id:
        q = q.filter(RentalBooking.id != exclude_id)
    return q.all()


def appointments_overlapping(start, end, camera_id, exclude_id=None):
    """Active appointments that reserve the SAME camera and overlap [start, end)."""
    if not camera_id:
        return []
    q = Appointment.query.filter(
        Appointment.camera_id == camera_id,
        Appointment.status != 'cancelled',
        Appointment.start_time < end,
        Appointment.end_time > start,
    )
    if exclude_id:
        q = q.filter(Appointment.id != exclude_id)
    return q.all()


def confirm_rental_or_rollback(b):
    """Race-safe finalize: after committing rental b, re-verify no *earlier* rental or any
    appointment reservation grabbed the same slot concurrently. Deterministic tiebreak — the
    smaller-id rental wins, and any unit-reserving appointment beats a rental. Returns an error
    string (and rolls b back) if b loses, else None."""
    earlier = [o for o in bookings_overlapping(b.start_date, b.end_date, b.camera_id, exclude_id=b.id) if o.id < b.id]
    appt_clash = appointments_overlapping(b.start_date, b.end_date, b.camera_id)
    if earlier or appt_clash:
        name = b.camera.name if b.camera else ''
        db.session.delete(b)
        db.session.commit()
        return f'Máy “{name}” vừa có người khác đặt trùng khung giờ. Vui lòng chọn thời gian khác.'
    return None


# ── Admin custom views (live inside the Flask-Admin panel) ─────────────────────

class BookingGridView(BaseView):
    """Google-Sheets-style grid: rows = cameras, columns = days, cells = bookings."""

    @expose('/')
    def index(self):
        cameras   = rent_cameras_ordered()
        color_map = color_map_for(cameras)
        # All active cameras selectable when an appointment reserves a specific unit.
        appt_cameras = (Camera.query.filter(Camera.is_sold.isnot(True))
                        .order_by(Camera.type, Camera.brand, Camera.name).all())
        return self.render('admin/booking_grid.html', cameras=cameras,
                           color_map=color_map, appt_cameras=appt_cameras)

    @expose('/data')
    def data(self):
        cameras   = rent_cameras_ordered()
        color_map = color_map_for(cameras)
        start = parse_dt(request.args.get('start')) or datetime.now()
        start = start.replace(hour=0, minute=0, second=0, microsecond=0)   # align to day columns
        try:
            days = max(1, min(60, int(request.args.get('days', 14))))
        except (TypeError, ValueError):
            days = 14
        end = start + timedelta(days=days)
        bookings = (RentalBooking.query
                    .filter(RentalBooking.start_date < end, RentalBooking.end_date > start)
                    .all())
        appts = (Appointment.query
                 .filter(Appointment.status != 'cancelled',
                         Appointment.start_time >= start, Appointment.start_time < end)
                 .order_by(Appointment.start_time)
                 .all())
        return jsonify({
            'cameras':  [{'id': c.id, 'name': c.name, 'brand': c.brand,
                          'color': color_map.get(c.id, '#6b7280')} for c in cameras],
            'bookings': [serialize_booking(b, color_map.get(b.camera_id)) for b in bookings],
            'appointments': [serialize_appt(a) for a in appts],
        })

    @expose('/save', methods=['POST'])
    def save(self):
        d = request.get_json(silent=True) or {}
        cam = db.session.get(Camera,d.get('camera_id'))
        start = parse_dt(d.get('start_date'))
        end   = parse_dt(d.get('end_date'))
        if not cam:
            return jsonify({'ok': False, 'error': 'Không tìm thấy máy ảnh'}), 400
        if not start or not end or end <= start:
            return jsonify({'ok': False, 'error': 'Ngày nhận / trả không hợp lệ'}), 400

        bid = d.get('id')
        # Block double-booking the same camera for an overlapping time (exclude self on edit).
        overlaps = bookings_overlapping(start, end, cam.id, exclude_id=int(bid) if bid else None)
        if overlaps:
            names = ', '.join(o.customer_name or '?' for o in overlaps)
            return jsonify({'ok': False,
                            'error': f'Máy “{cam.name}” đã được đặt trùng khung giờ này (khách: {names}). Chọn thời gian khác.'}), 409
        # Block if a buy-appointment has reserved this exact unit for the same window.
        appt_clash = appointments_overlapping(start, end, cam.id)
        if appt_clash:
            who = ', '.join(a.customer_name or '?' for a in appt_clash)
            return jsonify({'ok': False,
                            'error': f'Máy “{cam.name}” đang được giữ cho lịch hẹn mua (khách: {who}). Chọn thời gian khác.'}), 409

        if bid:
            b = db.session.get(RentalBooking,bid)
            if not b:
                return jsonify({'ok': False, 'error': 'Không tìm thấy đơn'}), 404
        else:
            b = RentalBooking()
            db.session.add(b)

        b.camera_id     = cam.id
        b.customer_name = (d.get('customer_name') or '').strip()
        b.phone         = (d.get('phone') or '').strip()
        b.notes         = (d.get('notes') or '').strip()
        b.start_date    = start
        b.end_date      = end
        b.total_price   = compute_rental_total(cam, start, end)
        db.session.commit()

        race_err = confirm_rental_or_rollback(b)
        if race_err:
            return jsonify({'ok': False, 'error': race_err}), 409

        cmap = color_map_for(rent_cameras_ordered())
        return jsonify({'ok': True, 'booking': serialize_booking(b, cmap.get(cam.id))})

    @expose('/delete', methods=['POST'])
    def delete(self):
        d = request.get_json(silent=True) or {}
        b = db.session.get(RentalBooking,d.get('id'))
        if not b:
            return jsonify({'ok': False, 'error': 'Không tìm thấy đơn'}), 404
        db.session.delete(b)
        db.session.commit()
        return jsonify({'ok': True})

    # ── Combined calendar feed: rentals + buy-appointments ───────────────────
    @expose('/events')
    def events(self):
        show = request.args.get('show', 'both')
        cameras   = rent_cameras_ordered()
        color_map = color_map_for(cameras)
        start = parse_dt((request.args.get('start') or '')[:19])
        end   = parse_dt((request.args.get('end') or '')[:19])
        events = []
        if show in ('both', 'rent'):
            q = RentalBooking.query
            if start: q = q.filter(RentalBooking.end_date >= start)
            if end:   q = q.filter(RentalBooking.start_date <= end)
            for b in q.all():
                color = color_map.get(b.camera_id, '#6b7280')
                cam = b.camera
                events.append({
                    'id': f'r{b.id}',
                    'title': f'{cam.name if cam else "Máy?"} · {b.customer_name}',
                    'start': b.start_date.isoformat(), 'end': b.end_date.isoformat(),
                    'backgroundColor': color, 'borderColor': color,
                    'extendedProps': serialize_booking(b, color) | {'kind': 'rent'},
                })
        if show in ('both', 'appt'):
            qa = Appointment.query.filter(Appointment.status != 'cancelled')
            if start: qa = qa.filter(Appointment.start_time >= start - timedelta(days=1))
            if end:   qa = qa.filter(Appointment.start_time <= end)
            for a in qa.all():
                title = '🛒 ' + (a.customer_name or 'Khách') + (f' · {a.interest}' if a.interest else '')
                events.append({
                    'id': f'a{a.id}', 'title': title,
                    'start': a.start_time.isoformat(),
                    'end': (a.end_time or (a.start_time + timedelta(minutes=45))).isoformat(),
                    'backgroundColor': APPT_COLOR, 'borderColor': APPT_COLOR,
                    'extendedProps': serialize_appt(a),
                })
        return jsonify(events)

    # ── Buy-appointment CRUD ─────────────────────────────────────────────────
    @expose('/appt/save', methods=['POST'])
    def appt_save(self):
        d = request.get_json(silent=True) or {}
        start = parse_dt(d.get('start_time'))
        if not start:
            return jsonify({'ok': False, 'error': 'Thời gian hẹn không hợp lệ'}), 400
        aid = d.get('id')
        end = parse_dt(d.get('end_time'))
        end = end if (end and end > start) else (start + timedelta(minutes=45))

        # Optional: reserve a specific unit. If set, block if that unit is rented or
        # already reserved by another appointment for an overlapping window.
        cam_id = d.get('camera_id') or None
        status = d.get('status') if d.get('status') in ('booked', 'done', 'cancelled') else 'booked'
        if cam_id and status != 'cancelled':
            cam = db.session.get(Camera, cam_id)
            if not cam:
                return jsonify({'ok': False, 'error': 'Không tìm thấy máy'}), 400
            if bookings_overlapping(start, end, cam.id):
                return jsonify({'ok': False, 'error': f'Máy “{cam.name}” đang có đơn thuê trùng khung giờ này.'}), 409
            appt_clash = appointments_overlapping(start, end, cam.id, exclude_id=int(aid) if aid else None)
            if appt_clash:
                who = ', '.join(x.customer_name or '?' for x in appt_clash)
                return jsonify({'ok': False, 'error': f'Máy “{cam.name}” đã được giữ cho lịch hẹn khác (khách: {who}).'}), 409

        if aid:
            a = db.session.get(Appointment, aid)
            if not a:
                return jsonify({'ok': False, 'error': 'Không tìm thấy lịch hẹn'}), 404
        else:
            a = Appointment()
            db.session.add(a)
        a.customer_name = (d.get('customer_name') or '').strip()
        a.phone         = (d.get('phone') or '').strip()
        a.interest      = (d.get('interest') or '').strip()
        a.notes         = (d.get('notes') or '').strip()
        a.start_time    = start
        a.end_time      = end
        a.camera_id     = int(cam_id) if cam_id else None
        a.status        = status
        db.session.commit()
        return jsonify({'ok': True, 'appt': serialize_appt(a)})

    @expose('/appt/delete', methods=['POST'])
    def appt_delete(self):
        d = request.get_json(silent=True) or {}
        a = db.session.get(Appointment, d.get('id'))
        if not a:
            return jsonify({'ok': False, 'error': 'Không tìm thấy lịch hẹn'}), 404
        db.session.delete(a)
        db.session.commit()
        return jsonify({'ok': True})


class CalendarView(BaseView):
    """FullCalendar week / 3-day time-grid view."""

    @expose('/')
    def index(self):
        cameras   = rent_cameras_ordered()
        color_map = color_map_for(cameras)
        return self.render('admin/calendar.html', cameras=cameras, color_map=color_map)

    @expose('/events')
    def events(self):
        cameras   = rent_cameras_ordered()
        color_map = color_map_for(cameras)
        start = parse_dt((request.args.get('start') or '')[:19])
        end   = parse_dt((request.args.get('end') or '')[:19])
        q = RentalBooking.query
        if start:
            q = q.filter(RentalBooking.end_date >= start)
        if end:
            q = q.filter(RentalBooking.start_date <= end)
        events = []
        for b in q.all():
            color = color_map.get(b.camera_id, '#6b7280')
            cam   = b.camera
            events.append({
                'id':              b.id,
                'title':           f'{cam.name if cam else "Máy?"} · {b.customer_name}',
                'start':           b.start_date.isoformat(),
                'end':             b.end_date.isoformat(),
                'backgroundColor': color,
                'borderColor':     color,
                'extendedProps':   serialize_booking(b, color),
            })
        return jsonify(events)


def serialize_variant(v):
    return {
        'id': v.id, 'camera_id': v.camera_id, 'color': v.color or '',
        'qty_available': v.qty_available or 0, 'qty_incoming': v.qty_incoming or 0,
        'qty_broken': v.qty_broken or 0, 'qty_unfixable': v.qty_unfixable or 0,
    }


def serialize_sale(s):
    return {
        'id': s.id, 'camera_id': s.camera_id, 'variant_id': s.variant_id,
        'color': s.color or '', 'quantity': s.quantity or 0, 'unit_price': s.unit_price or 0,
        'unit_cost': s.unit_cost or 0, 'discount': s.discount or 0,
        'total': s.line_total, 'note': s.note or '',
        'customer_name': s.customer_name or '', 'phone': s.phone or '',
        'payment_method': s.payment_method or 'cash', 'payment_status': s.payment_status or 'paid',
        'serial': s.serial or '', 'condition': s.condition or '',
        'warranty_until': s.warranty_until.strftime('%Y-%m-%d') if s.warranty_until else '',
        'is_return': bool(s.is_return), 'return_of': s.return_of,
        'margin': int(s.line_total - (s.quantity or 0) * (s.unit_cost or 0)),
        'date': s.sale_date.strftime('%d/%m/%Y') if s.sale_date else '',
    }


def serialize_receipt(r):
    return {
        'id': r.id, 'camera_id': r.camera_id, 'variant_id': r.variant_id,
        'supplier_id': r.supplier_id, 'supplier': r.supplier.name if r.supplier else '',
        'color': r.color or '', 'quantity': r.quantity or 0, 'unit_cost': r.unit_cost or 0,
        'total': (r.quantity or 0) * (r.unit_cost or 0), 'status': r.status or 'received',
        'note': r.note or '',
        'expected_date': r.expected_date.strftime('%Y-%m-%d') if r.expected_date else '',
        'date': (r.received_date or r.created_at).strftime('%d/%m/%Y') if (r.received_date or r.created_at) else '',
    }


def serialize_supplier(s):
    return {'id': s.id, 'name': s.name, 'phone': s.phone or '',
            'note': s.note or '', 'total_purchased': s.total_purchased}


def serialize_unit(cam):
    """A single for-sale camera unit (Selling ledger row)."""
    return {
        'id': cam.id, 'name': cam.name, 'brand': cam.brand or '',
        'origin': cam.origin or '', 'accessory': cam.accessory or '', 'color': cam.color or '',
        'category': cam.category or '',
        'date_in': cam.date_in.strftime('%Y-%m-%d') if cam.date_in else '',
        'note': cam.description or '', 'sold_price': cam.sold_price or 0,
        'import_cost': cam.import_cost or 0, 'repair_cost': cam.repair_cost or 0, 'profit': cam.profit,
        'sale_state': cam.sale_state or 'stock',
        'sold_to': cam.sold_to or '', 'sold_phone': cam.sold_phone or '', 'sold_source': cam.sold_source or '',
        'sold_at': cam.sold_date.strftime('%Y-%m-%dT%H:%M') if cam.sold_date else '',
        'sold_month': cam.sold_date.strftime('%Y-%m') if cam.sold_date else '',
        'gift_cost': cam.gift_cost, 'gift_label': cam.gift_label,
    }


def serialize_cost(c):
    return {'id': c.id, 'category': c.category or 'Khác', 'note': c.note or '',
            'amount': c.amount or 0,
            'date': c.cost_date.strftime('%Y-%m-%d') if c.cost_date else ''}


def serialize_accessory(a):
    return {'id': a.id, 'name': a.name, 'category': a.category or '',
            'cost': a.cost or 0, 'price': a.price or 0, 'stock': a.stock or 0,
            'note': a.note or '', 'stock_value': a.stock_value}


def camera_pnl(cam):
    return {
        'units_sold': cam.units_sold, 'rental_revenue': cam.rental_revenue,
        'sale_revenue': cam.sale_revenue, 'revenue': cam.revenue,
        'cost': cam.cost, 'profit': cam.profit, 'inventory_value': cam.inventory_value,
        'sale_state': cam.sale_state, 'state_label': cam.state_label,
    }


def camera_stock(cam):
    return {'available': cam.total_available, 'incoming': cam.total_incoming,
            'broken': cam.total_broken, 'unfixable': cam.total_unfixable}


class CamerasView(BaseView):
    """Inventory manager. Renting tab = editable table; Selling tab = per-color/status stock."""

    # Whitelist of editable fields → caster (prevents mass-assignment).
    EDITABLE = {
        'name':        str,   'brand':       str,   'badge':       str,   'type': str,
        'category':    str,   'description': str,   'color':       str,
        'origin':      str,   'accessory':   str,   'sold_to':     str,   'sold_phone': str,
        'sold_source': str,
        'price':       int,   'price_d2':    int,   'price_d3':    int,  'price_d4': int,
        'stock':       int,   'import_cost': int,   'repair_cost': int,  'sold_price': int,
        'reorder_point': int,
        'is_broken':   bool,  'is_sold':     bool,  'featured':    bool,
    }
    SALE_STATES = ('stock', 'sold', 'fixing', 'unfixable')

    @expose('/')
    def index(self):
        rent = Camera.query.filter_by(type='Rent').order_by(Camera.brand, Camera.name).all()
        # Active selling inventory = everything NOT yet sold (stock / fixing / unfixable).
        sale = (Camera.query.filter_by(type='Sale')
                .filter(Camera.sale_state != 'sold').order_by(Camera.id).all())
        # Sold units live in their own tab, newest sale first.
        sold = (Camera.query.filter_by(type='Sale', sale_state='sold')
                .order_by(Camera.sold_date.desc(), Camera.id.desc()).all())
        color_map = color_map_for(rent_cameras_ordered())   # rental colors match the grid
        brands = sorted({c.brand for c in (rent + sale + sold) if c.brand})
        gift_options = (Accessory.query.filter(Accessory.stock > 0)
                        .order_by(Accessory.category, Accessory.name).all())
        return self.render('admin/cameras.html',
                           rent_cameras=rent, sale_cameras=sale, sold_cameras=sold,
                           cameras=rent + sale + sold, color_map=color_map,
                           categories=CAMERA_CATEGORIES, brands=brands,
                           sale_sources=SALE_SOURCES,
                           gift_options=[serialize_accessory(a) for a in gift_options])

    @expose('/update', methods=['POST'])
    def update(self):
        d = request.get_json(silent=True) or {}
        cam = db.session.get(Camera, d.get('id'))
        field, value = d.get('field'), d.get('value')
        if not cam:
            return jsonify({'ok': False, 'error': 'Không tìm thấy máy'}), 404

        # ── per-unit special fields ──
        if field == 'date_in':
            cam.date_in = parse_dt(value) if value else None
            db.session.commit()
            return jsonify({'ok': True, 'pnl': camera_pnl(cam)})
        if field == 'sold_date':
            cam.sold_date = parse_dt(value) if value else None
            db.session.commit()
            return jsonify({'ok': True, 'pnl': camera_pnl(cam)})
        if field == 'sale_state':
            if value not in self.SALE_STATES:
                return jsonify({'ok': False, 'error': 'Trạng thái không hợp lệ'}), 400
            cam.sale_state = value
            cam.is_sold = (value == 'sold')
            if value == 'sold':
                cam.sold_date = datetime.now()
            else:                                  # reverting out of "sold" clears the sale
                cam.sold_date = None
                if value == 'stock':
                    # put any gifted accessories back into stock
                    for g in cam.gifts:
                        acc = db.session.get(Accessory, g.get('id'))
                        if acc:
                            acc.stock = (acc.stock or 0) + 1
                    cam.gift_json = ''
                    cam.sold_price = 0
                    cam.sold_to = cam.sold_phone = cam.sold_source = ''
            db.session.commit()
            return jsonify({'ok': True, 'pnl': camera_pnl(cam)})

        if field not in self.EDITABLE:
            return jsonify({'ok': False, 'error': 'Trường không hợp lệ'}), 400
        caster = self.EDITABLE[field]
        try:
            if caster is bool:
                cast = bool(value) if isinstance(value, bool) else str(value).lower() in ('1', 'true', 'on', 'yes')
            elif caster is int:
                cast = int(value or 0)
            else:
                cast = (str(value) if value is not None else '').strip()
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Giá trị không hợp lệ'}), 400
        if field == 'type' and cast not in ('Rent', 'Sale'):
            return jsonify({'ok': False, 'error': 'Loại máy không hợp lệ'}), 400
        setattr(cam, field, cast)
        if field == 'is_sold':
            cam.sold_date = datetime.now() if cast else None
        db.session.commit()
        return jsonify({'ok': True, 'pnl': camera_pnl(cam)})

    @expose('/sell', methods=['POST'])
    def sell(self):
        """Mark a for-sale unit as sold, capturing price + customer info."""
        d = request.get_json(silent=True) or {}
        cam = db.session.get(Camera, d.get('id'))
        if not cam:
            return jsonify({'ok': False, 'error': 'Không tìm thấy máy'}), 404
        try:
            price = max(0, int(d.get('sold_price') or 0))
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Giá bán không hợp lệ'}), 400
        cam.sale_state = 'sold'
        cam.is_sold = True
        cam.sold_price = price
        cam.sold_to = (d.get('customer_name') or '').strip()
        cam.sold_phone = (d.get('phone') or '').strip()
        cam.sold_source = (d.get('source') or '').strip()
        cam.sold_date = datetime.now()
        # gifted accessories: deduct each from stock, record cost (eats profit)
        gift_ids = d.get('gifts') or []
        gifts = []
        for aid in gift_ids:
            acc = db.session.get(Accessory, aid)
            if acc and (acc.stock or 0) > 0:
                acc.stock -= 1
                gifts.append({'id': acc.id, 'name': acc.name, 'cost': int(acc.cost or 0)})
        cam.gift_json = json.dumps(gifts, ensure_ascii=False) if gifts else ''
        db.session.commit()
        return jsonify({'ok': True, 'pnl': camera_pnl(cam), 'unit': serialize_unit(cam)})

    @expose('/add', methods=['POST'])
    def add(self):
        d = request.get_json(silent=True) or {}
        name = (d.get('name') or '').strip()
        if not name:
            return jsonify({'ok': False, 'error': 'Tên máy bắt buộc'}), 400
        base = ''.join(ch.lower() if ch.isalnum() else '-' for ch in name).strip('-')
        slug, n = base or 'may', 1
        while Camera.query.filter_by(slug=slug).first():
            n += 1
            slug = f'{base}-{n}'
        cam_type = d.get('type') if d.get('type') in ('Rent', 'Sale') else 'Rent'
        category = (d.get('category') or '').strip()
        cam = Camera(name=name, slug=slug, brand=(d.get('brand') or '').strip(),
                     type=cam_type, category=category, price=int(d.get('price') or 0),
                     import_cost=int(d.get('import_cost') or 0), stock=1,
                     origin=(d.get('origin') or '').strip(),
                     accessory=(d.get('accessory') or '').strip(),
                     description=(d.get('note') or '').strip(),
                     sale_state='stock')
        if d.get('date_in'):
            cam.date_in = parse_dt(d.get('date_in'))
        elif cam_type == 'Sale':
            cam.date_in = datetime.now()
        db.session.add(cam)
        db.session.commit()
        return jsonify({'ok': True, 'id': cam.id, 'camera': serialize_unit(cam)})

    # ── Excel import (bulk) + downloadable template ──────────────────────────
    @expose('/import-template')
    def import_template(self):
        """Download a blank .xlsx the shop can fill in and re-upload."""
        if not HAS_OPENPYXL:
            return jsonify({'ok': False, 'error': 'Thiếu thư viện openpyxl trên máy chủ.'}), 500
        wb = Workbook()
        ws = wb.active
        ws.title = 'Nhập máy'
        head_fill = PatternFill('solid', fgColor='B91C1C')
        head_font = Font(bold=True, color='FFFFFF')
        for ci, title in enumerate(IMPORT_COLUMNS, start=1):
            cell = ws.cell(row=1, column=ci, value=title)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[cell.column_letter].width = max(14, len(title) + 4)
        # one example row so the format is obvious
        ws.append(['LUMIX FX60', 'Bạc', 'Compact', 'N', 'Flash, pin',
                   '2026-06-01', 'Máy đẹp', 1750000, '', 'Còn hàng'])
        ws.freeze_panes = 'A2'
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='mau-nhap-may.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    @expose('/import-excel', methods=['POST'])
    def import_excel(self):
        """Bulk-create for-sale units from an uploaded .xlsx (same columns as the template)."""
        if not HAS_OPENPYXL:
            return jsonify({'ok': False, 'error': 'Thiếu thư viện openpyxl trên máy chủ.'}), 500
        f = request.files.get('file')
        if not f or not f.filename:
            return jsonify({'ok': False, 'error': 'Chưa chọn file.'}), 400
        try:
            wb = load_workbook(io.BytesIO(f.read()), data_only=True)
        except Exception:
            return jsonify({'ok': False, 'error': 'Không đọc được file Excel (.xlsx).'}), 400
        ws = wb.active
        added, errors = 0, []
        rows = list(ws.iter_rows(values_only=True))
        for ri, row in enumerate(rows, start=1):
            if ri == 1:
                continue                                   # header
            cells = list(row) + [None] * (len(IMPORT_COLUMNS) - len(row))
            name = (str(cells[0]).strip() if cells[0] is not None else '')
            if not name:
                continue                                   # skip blank lines
            try:
                base = ''.join(ch.lower() if ch.isalnum() else '-' for ch in name).strip('-')
                slug, n = base or 'may', 1
                while Camera.query.filter_by(slug=slug).first():
                    n += 1
                    slug = f'{base}-{n}'
                state = normalize_state(cells[9])
                sold_price = _to_int(cells[8])
                date_in = _cell_date(cells[5]) or datetime.now()
                cam = Camera(
                    name=name, slug=slug, type='Sale',
                    color=(str(cells[1]).strip() if cells[1] else ''),
                    category=(str(cells[2]).strip() if cells[2] else ''),
                    origin=normalize_origin(cells[3]),
                    accessory=(str(cells[4]).strip() if cells[4] else ''),
                    date_in=date_in,
                    description=(str(cells[6]).strip() if cells[6] else ''),
                    import_cost=_to_int(cells[7]),
                    sold_price=sold_price,
                    sale_state=state, price=sold_price or 0, stock=1)
                if state == 'sold':
                    cam.is_sold = True
                    cam.sold_date = date_in
                db.session.add(cam)
                added += 1
            except Exception as e:                         # noqa: BLE001
                errors.append(f'Dòng {ri}: {e}')
        db.session.commit()
        return jsonify({'ok': True, 'added': added, 'errors': errors[:10]})

    @expose('/delete', methods=['POST'])
    def delete(self):
        d = request.get_json(silent=True) or {}
        cam = db.session.get(Camera, d.get('id'))
        if not cam:
            return jsonify({'ok': False, 'error': 'Không tìm thấy máy'}), 404
        RentalBooking.query.filter_by(camera_id=cam.id).delete()
        db.session.delete(cam)   # variants cascade via relationship
        db.session.commit()
        return jsonify({'ok': True})

    # ── Color/status variants (Selling tab) ──────────────────────────────────
    @expose('/variant/save', methods=['POST'])
    def variant_save(self):
        d = request.get_json(silent=True) or {}
        cam = db.session.get(Camera, d.get('camera_id'))
        if not cam:
            return jsonify({'ok': False, 'error': 'Không tìm thấy máy'}), 404
        vid = d.get('id')
        if vid:
            v = db.session.get(CameraVariant, vid)
            if not v or v.camera_id != cam.id:
                return jsonify({'ok': False, 'error': 'Không tìm thấy màu'}), 404
        else:
            v = CameraVariant(camera_id=cam.id)
            db.session.add(v)
        v.color = (d.get('color') or '').strip()
        for q in self.VARIANT_QTY:
            try:
                setattr(v, q, max(0, int(d.get(q) or 0)))
            except (TypeError, ValueError):
                setattr(v, q, 0)
        db.session.commit()
        return jsonify({'ok': True, 'variant': serialize_variant(v),
                        'totals': camera_stock(cam), 'pnl': camera_pnl(cam)})

    @expose('/variant/delete', methods=['POST'])
    def variant_delete(self):
        d = request.get_json(silent=True) or {}
        v = db.session.get(CameraVariant, d.get('id'))
        if not v:
            return jsonify({'ok': False, 'error': 'Không tìm thấy màu'}), 404
        cam = v.camera
        db.session.delete(v)
        db.session.commit()
        return jsonify({'ok': True, 'totals': camera_stock(cam), 'pnl': camera_pnl(cam)})

    # ── Receiving / import batches (logs cost → weighted-average COGS) ────────
    @expose('/receive', methods=['POST'])
    def receive(self):
        d = request.get_json(silent=True) or {}
        cam = db.session.get(Camera, d.get('camera_id'))
        if not cam:
            return jsonify({'ok': False, 'error': 'Không tìm thấy máy'}), 404
        try:
            qty       = max(1, int(d.get('quantity') or 0))
            unit_cost = max(0, int(d.get('unit_cost') or 0))
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Số lượng / giá nhập không hợp lệ'}), 400
        status = 'ordered' if d.get('status') == 'ordered' else 'received'
        color  = (d.get('color') or '').strip()

        # locate or create the color variant
        variant = None
        vid = d.get('variant_id')
        if vid:
            variant = db.session.get(CameraVariant, vid)
            if variant and variant.camera_id != cam.id:
                variant = None
        if not variant and color:
            variant = next((v for v in cam.variants
                            if (v.color or '').strip().lower() == color.lower()), None)
        if not variant:
            variant = CameraVariant(camera_id=cam.id, color=color)
            db.session.add(variant)
            db.session.flush()

        supplier = db.session.get(Supplier, d.get('supplier_id')) if d.get('supplier_id') else None
        rec = StockReceipt(
            camera_id=cam.id, variant_id=variant.id, color=variant.color or color,
            supplier_id=(supplier.id if supplier else None),
            quantity=qty, unit_cost=unit_cost, status=status,
            expected_date=parse_dt(d.get('expected_date')),
            received_date=(datetime.now() if status == 'received' else None),
            note=(d.get('note') or '').strip())
        db.session.add(rec)
        if status == 'received':
            variant.qty_available = (variant.qty_available or 0) + qty
        else:
            variant.qty_incoming = (variant.qty_incoming or 0) + qty
        db.session.commit()
        return jsonify({'ok': True, 'receipt': serialize_receipt(rec),
                        'variant': serialize_variant(variant),
                        'totals': camera_stock(cam), 'pnl': camera_pnl(cam)})

    @expose('/receipt/mark-received', methods=['POST'])
    def receipt_mark_received(self):
        d = request.get_json(silent=True) or {}
        rec = db.session.get(StockReceipt, d.get('id'))
        if not rec:
            return jsonify({'ok': False, 'error': 'Không tìm thấy phiếu nhập'}), 404
        if rec.status != 'received':
            rec.status = 'received'
            rec.received_date = datetime.now()
            variant = db.session.get(CameraVariant, rec.variant_id) if rec.variant_id else None
            if variant:
                variant.qty_incoming  = max(0, (variant.qty_incoming or 0) - (rec.quantity or 0))
                variant.qty_available = (variant.qty_available or 0) + (rec.quantity or 0)
            db.session.commit()
        cam = rec.camera
        return jsonify({'ok': True, 'receipt': serialize_receipt(rec),
                        'totals': camera_stock(cam), 'pnl': camera_pnl(cam)})

    # ── Move stock between statuses (đã về / hỏng / không sửa được) ───────────
    @expose('/variant/move', methods=['POST'])
    def variant_move(self):
        d = request.get_json(silent=True) or {}
        v = db.session.get(CameraVariant, d.get('id'))
        if not v:
            return jsonify({'ok': False, 'error': 'Không tìm thấy màu'}), 404
        src, dst = d.get('from'), d.get('to')
        if src not in self.VARIANT_QTY or dst not in self.VARIANT_QTY:
            return jsonify({'ok': False, 'error': 'Trạng thái không hợp lệ'}), 400
        try:
            qty = max(1, int(d.get('quantity') or 1))
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Số lượng không hợp lệ'}), 400
        if (getattr(v, src) or 0) < qty:
            return jsonify({'ok': False, 'error': 'Không đủ số lượng để chuyển'}), 400
        setattr(v, src, (getattr(v, src) or 0) - qty)
        setattr(v, dst, (getattr(v, dst) or 0) + qty)
        cam = v.camera
        db.session.commit()
        return jsonify({'ok': True, 'variant': serialize_variant(v),
                        'totals': camera_stock(cam), 'pnl': camera_pnl(cam)})

    # ── Suppliers ─────────────────────────────────────────────────────────────
    @expose('/supplier/add', methods=['POST'])
    def supplier_add(self):
        d = request.get_json(silent=True) or {}
        name = (d.get('name') or '').strip()
        if not name:
            return jsonify({'ok': False, 'error': 'Tên nhà cung cấp bắt buộc'}), 400
        s = Supplier(name=name, phone=(d.get('phone') or '').strip(), note=(d.get('note') or '').strip())
        db.session.add(s)
        db.session.commit()
        return jsonify({'ok': True, 'supplier': serialize_supplier(s)})

    # ── Sales log (Selling tab) ──────────────────────────────────────────────
    @expose('/sale/add', methods=['POST'])
    def sale_add(self):
        d = request.get_json(silent=True) or {}
        cam = db.session.get(Camera, d.get('camera_id'))
        if not cam:
            return jsonify({'ok': False, 'error': 'Không tìm thấy máy'}), 404
        try:
            qty      = max(1, int(d.get('quantity') or 1))
            price    = max(0, int(d.get('unit_price') or 0))
            discount = max(0, int(d.get('discount') or 0))
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Số lượng / đơn giá không hợp lệ'}), 400

        color, variant = (d.get('color') or '').strip(), None
        vid = d.get('variant_id')
        if vid:
            variant = db.session.get(CameraVariant, vid)
            if variant and variant.camera_id == cam.id:
                if qty > (variant.qty_available or 0):       # oversell guard
                    return jsonify({'ok': False, 'error':
                        f'Chỉ còn {variant.qty_available or 0} máy màu “{variant.color or color}” — không đủ để bán {qty}.'}), 400
                color = variant.color or color
                variant.qty_available = (variant.qty_available or 0) - qty
            else:
                variant = None

        line_total = qty * price - discount
        status = (d.get('payment_status') or 'paid').strip()
        if status == 'paid':
            paid = line_total
        elif status == 'unpaid':
            paid = 0
        else:                                                 # deposit / installment
            paid = max(0, int(d.get('amount_paid') or 0))

        rec = SaleRecord(
            camera_id=cam.id, variant_id=(variant.id if variant else None),
            color=color, quantity=qty, unit_price=price, discount=discount,
            unit_cost=int(cam.avg_cost or 0),                 # immutable COGS snapshot
            note=(d.get('note') or '').strip(),
            customer_name=(d.get('customer_name') or '').strip(),
            phone=(d.get('phone') or '').strip(),
            payment_method=(d.get('payment_method') or 'cash').strip(),
            payment_status=status, amount_paid=paid,
            serial=(d.get('serial') or '').strip(),
            condition=(d.get('condition') or '').strip(),
            warranty_until=parse_dt(d.get('warranty_until')))
        db.session.add(rec)
        db.session.commit()
        return jsonify({'ok': True, 'sale': serialize_sale(rec), 'pnl': camera_pnl(cam),
                        'totals': camera_stock(cam),
                        'variant': ({'id': variant.id, 'qty_available': variant.qty_available} if variant else None)})

    @expose('/sale/return', methods=['POST'])
    def sale_return(self):
        d = request.get_json(silent=True) or {}
        orig = db.session.get(SaleRecord, d.get('id'))
        if not orig:
            return jsonify({'ok': False, 'error': 'Không tìm thấy đơn bán'}), 404
        if orig.is_return:
            return jsonify({'ok': False, 'error': 'Đây đã là phiếu trả hàng'}), 400
        cam = orig.camera
        try:
            qty = max(1, min(int(d.get('quantity') or orig.quantity or 1), orig.quantity or 1))
        except (TypeError, ValueError):
            qty = orig.quantity or 1
        ret = SaleRecord(
            camera_id=cam.id, variant_id=orig.variant_id, color=orig.color,
            quantity=qty, unit_price=orig.unit_price, discount=0,
            unit_cost=orig.unit_cost, is_return=True, return_of=orig.id,
            customer_name=orig.customer_name, phone=orig.phone,
            payment_method=orig.payment_method, payment_status='paid',
            note=(d.get('note') or 'Khách trả hàng').strip())
        db.session.add(ret)
        variant_info = None
        if orig.variant_id:
            variant = db.session.get(CameraVariant, orig.variant_id)
            if variant:
                variant.qty_available = (variant.qty_available or 0) + qty      # restock
                variant_info = {'id': variant.id, 'qty_available': variant.qty_available}
        db.session.commit()
        return jsonify({'ok': True, 'sale': serialize_sale(ret), 'pnl': camera_pnl(cam),
                        'totals': camera_stock(cam), 'variant': variant_info})

    @expose('/sale/delete', methods=['POST'])
    def sale_delete(self):
        d = request.get_json(silent=True) or {}
        rec = db.session.get(SaleRecord, d.get('id'))
        if not rec:
            return jsonify({'ok': False, 'error': 'Không tìm thấy đơn'}), 404
        cam = rec.camera
        variant_info = None
        if rec.variant_id:
            variant = db.session.get(CameraVariant, rec.variant_id)
            if variant:
                # reverse the stock effect: a sale removed qty (add back); a return added qty (remove)
                delta = -(rec.quantity or 0) if rec.is_return else (rec.quantity or 0)
                variant.qty_available = max(0, (variant.qty_available or 0) + delta)
                variant_info = {'id': variant.id, 'qty_available': variant.qty_available}
        db.session.delete(rec)
        db.session.commit()
        return jsonify({'ok': True, 'pnl': camera_pnl(cam),
                        'totals': camera_stock(cam), 'variant': variant_info})


class FinanceView(BaseView):
    """Revenue / cost / profit dashboard."""

    @staticmethod
    def _bucket(period, dt):
        """[start, end) of the period that `dt` falls into."""
        if period == 'day':
            s = dt.replace(hour=0, minute=0, second=0, microsecond=0)
            return s, s + timedelta(days=1)
        if period == 'year':
            return datetime(dt.year, 1, 1), datetime(dt.year + 1, 1, 1)
        end_y, end_m = (dt.year + 1, 1) if dt.month == 12 else (dt.year, dt.month + 1)
        return datetime(dt.year, dt.month, 1), datetime(end_y, end_m, 1)

    @staticmethod
    def _shift(period, dt, n):
        """Move the anchor date back `n` whole periods."""
        if period == 'day':
            return dt - timedelta(days=n)
        if period == 'year':
            return dt.replace(year=dt.year - n)
        total = (dt.year * 12 + (dt.month - 1)) - n
        return datetime(total // 12, total % 12 + 1, 1)

    @expose('/')
    def index(self):
        period = request.args.get('period', 'month')
        if period not in ('day', 'month', 'year'):
            period = 'month'
        now = datetime.now()

        # ── anchor: the specific day / month / year being viewed ──
        anchor_raw = (request.args.get('anchor') or '').strip()
        anchor = parse_dt(anchor_raw)
        if not anchor:
            if period == 'year' and anchor_raw.isdigit():
                anchor = datetime(int(anchor_raw), 1, 1)
            elif period == 'month' and len(anchor_raw) >= 7 and anchor_raw[:4].isdigit():
                try:
                    anchor = datetime(int(anchor_raw[:4]), int(anchor_raw[5:7]), 1)
                except ValueError:
                    anchor = None
        if not anchor:
            anchor = now
        astart, aend = self._bucket(period, anchor)

        cameras = Camera.query.order_by(Camera.brand, Camera.name).all()
        rent_cameras = [c for c in cameras if c.type == 'Rent']
        sale_cameras = [c for c in cameras if c.type == 'Sale']
        bookings  = RentalBooking.query.all()
        sold_cams = [c for c in sale_cameras if c.sale_state == 'sold' and c.sold_date]
        store_costs = StoreCost.query.all()
        acc_sales = AccessorySale.query.all()

        def window(start, end):
            """All money flows whose own date falls inside [start, end)."""
            rrev  = sum(int(b.total_price or 0) for b in bookings if b.start_date and start <= b.start_date < end)
            sold  = [c for c in sold_cams if start <= c.sold_date < end]
            srev  = sum(int(c.sold_price or 0) for c in sold)
            scogs = sum(int(c.cost) for c in sold)
            scost = sum(int(x.amount or 0) for x in store_costs if x.cost_date and start <= x.cost_date < end)
            asales = [s for s in acc_sales if s.sale_date and start <= s.sale_date < end]
            arev  = sum(s.revenue for s in asales)
            acogs = sum(s.cost_total for s in asales)
            return {'rrev': rrev, 'srev': srev, 'scogs': scogs, 'sprof': srev - scogs,
                    'scost': scost, 'sold': sold,
                    'arev': arev, 'acogs': acogs, 'aprof': arev - acogs}

        # ── trailing buckets for the trend charts (ending at the anchor) ──
        counts = {'day': 14, 'month': 12, 'year': 5}[period]
        buckets = []
        for i in range(counts - 1, -1, -1):
            bdt = self._shift(period, anchor, i)
            bs, be = self._bucket(period, bdt)
            label = bs.strftime('%d/%m') if period == 'day' else (str(bs.year) if period == 'year' else f'{bs.month:02d}/{bs.year}')
            buckets.append((label, bs, be, bs == astart))
        rental_series, sale_rev_series, sale_profit_series, store_cost_series = [], [], [], []
        for (label, bs, be, sel) in buckets:
            w = window(bs, be)
            rental_series.append({'label': label, 'value': int(w['rrev']), 'sel': sel})
            sale_rev_series.append({'label': label, 'value': int(w['srev']), 'sel': sel})
            sale_profit_series.append({'label': label, 'value': int(w['sprof']), 'sel': sel})
            store_cost_series.append({'label': label, 'value': int(w['scost']), 'sel': sel})

        # ── totals for the SELECTED period only (dates now matter) ──
        w = window(astart, aend)
        rental_revenue = int(w['rrev'])
        rental_cost    = 0                                # rental gear cost is not dated
        rental_profit  = rental_revenue - rental_cost
        sale_revenue   = int(w['srev'])
        sale_cost      = int(w['scogs'])
        sale_profit    = int(w['sprof'])
        accessory_revenue = int(w['arev'])
        accessory_cost    = int(w['acogs'])
        accessory_profit  = int(w['aprof'])
        store_cost_total = int(w['scost'])

        total_revenue = rental_revenue + sale_revenue + accessory_revenue
        total_profit  = rental_profit + sale_profit + accessory_profit
        total_cost    = rental_cost + sale_cost + accessory_cost
        net_profit    = total_profit - store_cost_total

        # ── inventory snapshot (current state, not date-scoped) ──
        in_stock     = [c for c in sale_cameras if c.sale_state == 'stock']
        fixing_units = [c for c in sale_cameras if c.sale_state == 'fixing']
        unfix_units  = [c for c in sale_cameras if c.sale_state == 'unfixable']
        inventory_value = int(sum(c.inventory_value for c in sale_cameras))
        sold_in_period = w['sold']
        top_profit = sorted(sold_in_period, key=lambda c: c.profit, reverse=True)[:8]
        aging = [{'cam': c, 'days': (now - c.date_in).days if c.date_in else None} for c in in_stock]
        aging.sort(key=lambda x: (-1 if x['days'] is None else x['days']), reverse=True)
        aging = aging[:8]

        summary = {
            'rental_revenue': rental_revenue, 'rental_cost': rental_cost, 'rental_profit': rental_profit,
            'sale_revenue': sale_revenue, 'sale_cost': sale_cost, 'sale_profit': sale_profit,
            'accessory_revenue': accessory_revenue, 'accessory_cost': accessory_cost, 'accessory_profit': accessory_profit,
            'total_revenue': total_revenue, 'total_cost': total_cost, 'total_profit': total_profit,
            'inventory_value': inventory_value,
            'store_cost_total': store_cost_total, 'net_profit': net_profit,
            'units_sold': len(sold_in_period), 'in_stock_count': len(in_stock),
            'broken_count': len(fixing_units) + len(unfix_units), 'unfixable_count': len(unfix_units),
            'sold_count': len(sold_in_period), 'camera_count': len(cameras),
        }
        period_labels = {'day': '14 ngày', 'month': '12 tháng', 'year': '5 năm'}
        anchor_labels = {
            'day':   'Ngày ' + astart.strftime('%d/%m/%Y'),
            'month': 'Tháng ' + astart.strftime('%m/%Y'),
            'year':  'Năm ' + astart.strftime('%Y'),
        }
        anchor_values = {'day': astart.strftime('%Y-%m-%d'),
                         'month': astart.strftime('%Y-%m'), 'year': astart.strftime('%Y')}
        prev_anchor = self._shift(period, anchor, 1)
        next_dt = self._shift(period, anchor, -1)
        prev_values = {'day': prev_anchor.strftime('%Y-%m-%d'),
                       'month': prev_anchor.strftime('%Y-%m'), 'year': prev_anchor.strftime('%Y')}
        next_values = {'day': next_dt.strftime('%Y-%m-%d'),
                       'month': next_dt.strftime('%Y-%m'), 'year': next_dt.strftime('%Y')}
        return self.render('admin/finance.html', summary=summary, cameras=cameras,
                           rent_cameras=rent_cameras, sale_cameras=sale_cameras,
                           period=period, period_label=period_labels[period],
                           anchor_label=anchor_labels[period], anchor_value=anchor_values[period],
                           prev_anchor=prev_values[period], next_anchor=next_values[period],
                           is_current=(astart <= now < aend),
                           rental_series=rental_series, sale_rev_series=sale_rev_series,
                           sale_profit_series=sale_profit_series, store_cost_series=store_cost_series,
                           top_profit=top_profit, aging=aging)


class StoreCostView(BaseView):
    """Shop overhead / operating costs (Chi phí tiệm) — rent, ads, shipping…"""

    CATEGORIES = ['Tiền nhà', 'Quảng cáo', 'Vận chuyển', 'Phụ kiện', 'Lương', 'Điện nước', 'Sửa chữa', 'Khác']

    @expose('/')
    def index(self):
        costs = StoreCost.query.order_by(StoreCost.cost_date.desc(), StoreCost.id.desc()).all()
        by_cat = {}
        for c in costs:
            by_cat[c.category or 'Khác'] = by_cat.get(c.category or 'Khác', 0) + (c.amount or 0)
        by_cat = sorted(by_cat.items(), key=lambda kv: kv[1], reverse=True)
        total = int(sum(c.amount or 0 for c in costs))
        now = datetime.now()
        month_total = int(sum((c.amount or 0) for c in costs
                              if c.cost_date and c.cost_date.year == now.year and c.cost_date.month == now.month))
        return self.render('admin/chiphi.html',
                           costs=[serialize_cost(c) for c in costs],
                           by_cat=by_cat, total=total, month_total=month_total,
                           categories=self.CATEGORIES, today=now.strftime('%Y-%m-%d'))

    @expose('/add', methods=['POST'])
    def add(self):
        d = request.get_json(silent=True) or {}
        try:
            amount = max(0, int(d.get('amount') or 0))
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Số tiền không hợp lệ'}), 400
        c = StoreCost(category=(d.get('category') or 'Khác').strip(),
                      note=(d.get('note') or '').strip(), amount=amount,
                      cost_date=parse_dt(d.get('cost_date')) or datetime.now())
        db.session.add(c)
        db.session.commit()
        return jsonify({'ok': True, 'cost': serialize_cost(c)})

    @expose('/update', methods=['POST'])
    def update(self):
        d = request.get_json(silent=True) or {}
        c = db.session.get(StoreCost, d.get('id'))
        if not c:
            return jsonify({'ok': False, 'error': 'Không tìm thấy'}), 404
        f, v = d.get('field'), d.get('value')
        if f == 'amount':
            try:
                c.amount = max(0, int(v or 0))
            except (TypeError, ValueError):
                return jsonify({'ok': False, 'error': 'Số tiền không hợp lệ'}), 400
        elif f == 'category':
            c.category = (str(v) if v else 'Khác').strip()
        elif f == 'note':
            c.note = (str(v) if v else '').strip()
        elif f == 'cost_date':
            c.cost_date = parse_dt(v) or c.cost_date
        else:
            return jsonify({'ok': False, 'error': 'Trường không hợp lệ'}), 400
        db.session.commit()
        return jsonify({'ok': True})

    @expose('/delete', methods=['POST'])
    def delete(self):
        d = request.get_json(silent=True) or {}
        c = db.session.get(StoreCost, d.get('id'))
        if not c:
            return jsonify({'ok': False, 'error': 'Không tìm thấy'}), 404
        db.session.delete(c)
        db.session.commit()
        return jsonify({'ok': True})


class AccessoriesView(BaseView):
    """Phụ kiện / hàng hoá khác — sold by quantity (lens, pin, thẻ nhớ…)."""

    EDITABLE = {'name': str, 'category': str, 'note': str,
                'cost': int, 'price': int, 'stock': int}

    @expose('/')
    def index(self):
        items = Accessory.query.order_by(Accessory.category, Accessory.name).all()
        total_stock  = int(sum(a.stock or 0 for a in items))
        stock_value  = int(sum(a.stock_value for a in items))
        now = datetime.now()
        month_sales = [s for s in AccessorySale.query.all()
                       if s.sale_date and s.sale_date.year == now.year and s.sale_date.month == now.month]
        month_rev = int(sum(s.revenue for s in month_sales))
        return self.render('admin/phukien.html',
                           accessories=[serialize_accessory(a) for a in items],
                           categories=ACCESSORY_CATEGORIES,
                           total_stock=total_stock, stock_value=stock_value,
                           item_count=len(items), month_rev=month_rev,
                           today=now.strftime('%Y-%m-%d'))

    @expose('/add', methods=['POST'])
    def add(self):
        d = request.get_json(silent=True) or {}
        name = (d.get('name') or '').strip()
        if not name:
            return jsonify({'ok': False, 'error': 'Tên phụ kiện bắt buộc'}), 400
        try:
            a = Accessory(name=name, category=(d.get('category') or '').strip(),
                          cost=max(0, int(d.get('cost') or 0)),
                          price=max(0, int(d.get('price') or 0)),
                          stock=max(0, int(d.get('stock') or 0)),
                          note=(d.get('note') or '').strip())
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Giá trị không hợp lệ'}), 400
        db.session.add(a)
        db.session.commit()
        return jsonify({'ok': True, 'accessory': serialize_accessory(a)})

    @expose('/update', methods=['POST'])
    def update(self):
        d = request.get_json(silent=True) or {}
        a = db.session.get(Accessory, d.get('id'))
        if not a:
            return jsonify({'ok': False, 'error': 'Không tìm thấy'}), 404
        field, value = d.get('field'), d.get('value')
        if field not in self.EDITABLE:
            return jsonify({'ok': False, 'error': 'Trường không hợp lệ'}), 400
        caster = self.EDITABLE[field]
        try:
            cast = max(0, int(value or 0)) if caster is int else (str(value) if value is not None else '').strip()
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Giá trị không hợp lệ'}), 400
        setattr(a, field, cast)
        db.session.commit()
        return jsonify({'ok': True, 'accessory': serialize_accessory(a)})

    @expose('/sell', methods=['POST'])
    def sell(self):
        """Sell N units at a price → decrement stock, log the sale."""
        d = request.get_json(silent=True) or {}
        a = db.session.get(Accessory, d.get('id'))
        if not a:
            return jsonify({'ok': False, 'error': 'Không tìm thấy'}), 404
        try:
            qty   = int(d.get('quantity') or 0)
            price = max(0, int(d.get('price') or 0))
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': 'Số lượng / giá không hợp lệ'}), 400
        if qty < 1:
            return jsonify({'ok': False, 'error': 'Số lượng phải ≥ 1'}), 400
        if qty > (a.stock or 0):
            return jsonify({'ok': False, 'error': f'Chỉ còn {a.stock or 0} cái trong kho'}), 400
        a.stock -= qty
        rec = AccessorySale(accessory_id=a.id, name=a.name, quantity=qty,
                            unit_price=price, unit_cost=int(a.cost or 0), sale_date=datetime.now())
        db.session.add(rec)
        db.session.commit()
        return jsonify({'ok': True, 'accessory': serialize_accessory(a),
                        'sold': qty, 'revenue': rec.revenue})

    @expose('/delete', methods=['POST'])
    def delete(self):
        d = request.get_json(silent=True) or {}
        a = db.session.get(Accessory, d.get('id'))
        if not a:
            return jsonify({'ok': False, 'error': 'Không tìm thấy'}), 404
        db.session.delete(a)
        db.session.commit()
        return jsonify({'ok': True})


class SheetView(BaseView):
    """A free-form Google-Sheets-style scratch sheet."""

    def _get_sheet(self):
        sheet = Sheet.query.first()
        if not sheet:
            blank = [['' for _ in range(8)] for _ in range(30)]
            sheet = Sheet(name='Sổ tay', data_json=json.dumps(blank))
            db.session.add(sheet)
            db.session.commit()
        return sheet

    @expose('/')
    def index(self):
        return self.render('admin/sheet.html')

    @expose('/load')
    def load(self):
        sheet = self._get_sheet()
        return jsonify({'name': sheet.name, 'data': sheet.data,
                        'updated_at': sheet.updated_at.isoformat() if sheet.updated_at else None})

    @expose('/save', methods=['POST'])
    def save(self):
        d = request.get_json(silent=True) or {}
        data = d.get('data')
        if not isinstance(data, list):
            return jsonify({'ok': False, 'error': 'Dữ liệu không hợp lệ'}), 400
        # Coerce to a clean 2D array of strings.
        clean = [[('' if c is None else str(c)) for c in (row if isinstance(row, list) else [])]
                 for row in data]
        sheet = self._get_sheet()
        if d.get('name'):
            sheet.name = str(d['name']).strip()[:100]
        sheet.data_json = json.dumps(clean, ensure_ascii=False)
        db.session.commit()
        return jsonify({'ok': True, 'updated_at': sheet.updated_at.isoformat() if sheet.updated_at else None})


admin.add_view(BookingGridView(name='Lịch thuê (Sheet)', endpoint='bookinggrid'))
admin.add_view(CalendarView(name='Lịch (Calendar)',     endpoint='calendar'))
admin.add_view(CamerasView(name='Quản lý máy',          endpoint='cameras'))
admin.add_view(AccessoriesView(name='Phụ kiện',         endpoint='phukien'))
admin.add_view(FinanceView(name='Tài chính',            endpoint='finance'))
admin.add_view(StoreCostView(name='Chi phí tiệm',       endpoint='chiphi'))
admin.add_view(SheetView(name='Sổ tay',                 endpoint='sheet'))


# ── Schema migration (non-destructive) ─────────────────────────────────────────

def _auto_add_missing_columns():
    """Generic safety net: for every mapped table that already exists, add any column
    present in the model but missing in the DB (as a nullable column — safe on SQLite,
    never drops or rewrites data). Catches anything the explicit lists above miss."""
    insp = sa_inspect(db.engine)
    existing = set(insp.get_table_names())
    for table in db.metadata.sorted_tables:
        if table.name not in existing:
            continue                      # brand-new tables are created in full by create_all
        have = {c['name'] for c in insp.get_columns(table.name)}
        for col in table.columns:
            if col.name in have:
                continue
            try:
                coltype = col.type.compile(dialect=db.engine.dialect)
            except Exception:
                coltype = 'TEXT'
            default = ''
            d = getattr(col, 'default', None)
            if d is not None and getattr(d, 'is_scalar', False):
                v = d.arg
                if isinstance(v, bool):            default = f' DEFAULT {1 if v else 0}'
                elif isinstance(v, (int, float)):  default = f' DEFAULT {v}'
                elif isinstance(v, str):           default = " DEFAULT '{}'".format(v.replace("'", "''"))
            db.session.execute(text(f'ALTER TABLE {table.name} ADD COLUMN {col.name} {coltype}{default}'))
    db.session.commit()


def ensure_schema():
    """Add any new columns / tables without dropping existing data."""
    insp = sa_inspect(db.engine)
    if 'camera' in insp.get_table_names():
        cols = {c['name'] for c in insp.get_columns('camera')}
        new_cols = {
            'import_cost':   'INTEGER DEFAULT 0',
            'is_broken':     'BOOLEAN DEFAULT 0',
            'repair_cost':   'INTEGER DEFAULT 0',
            'is_sold':       'BOOLEAN DEFAULT 0',
            'sold_price':    'INTEGER DEFAULT 0',
            'sold_date':     'DATETIME',
            'category':      "VARCHAR(40) DEFAULT ''",
            'reorder_point': 'INTEGER DEFAULT 0',
            'origin':        "VARCHAR(20) DEFAULT ''",
            'accessory':     "VARCHAR(120) DEFAULT ''",
            'color':         "VARCHAR(40) DEFAULT ''",
            'date_in':       'DATETIME',
            'sale_state':    "VARCHAR(20) DEFAULT 'stock'",
            'sold_to':       "VARCHAR(100) DEFAULT ''",
            'sold_phone':    "VARCHAR(20) DEFAULT ''",
            'sold_source':   "VARCHAR(30) DEFAULT ''",
            'gift_json':     "TEXT DEFAULT ''",
        }
        for name, ddl in new_cols.items():
            if name not in cols:
                db.session.execute(text(f'ALTER TABLE camera ADD COLUMN {name} {ddl}'))
        db.session.commit()

    if 'sale_record' in insp.get_table_names():
        scols = {c['name'] for c in insp.get_columns('sale_record')}
        sale_new = {
            'unit_cost':      'INTEGER DEFAULT 0',
            'discount':       'INTEGER DEFAULT 0',
            'customer_name':  "VARCHAR(100) DEFAULT ''",
            'phone':          "VARCHAR(20) DEFAULT ''",
            'payment_method': "VARCHAR(20) DEFAULT 'cash'",
            'payment_status': "VARCHAR(20) DEFAULT 'paid'",
            'amount_paid':    'INTEGER DEFAULT 0',
            'serial':         "VARCHAR(120) DEFAULT ''",
            'warranty_until': 'DATETIME',
            'condition':      "VARCHAR(20) DEFAULT ''",
            'is_return':      'BOOLEAN DEFAULT 0',
            'return_of':      'INTEGER',
        }
        for name, ddl in sale_new.items():
            if name not in scols:
                db.session.execute(text(f'ALTER TABLE sale_record ADD COLUMN {name} {ddl}'))
        db.session.commit()

    db.create_all()               # creates brand-new tables (Appointment, Sheet, Supplier…) in full
    _auto_add_missing_columns()   # backfill any column added to an already-existing table


# ── Auto-migrate on EVERY startup ──────────────────────────────────────────────
# Runs under Gunicorn too (the __main__ block below never executes there). This is
# non-destructive (only CREATE TABLE / ADD COLUMN) and never seeds, so the production
# database on the server keeps all its data and just gains the new schema.
with app.app_context():
    try:
        ensure_schema()
    except Exception as exc:      # never let a migration hiccup take the whole app down
        app.logger.warning(f'[schema] ensure_schema at startup failed: {exc}')


# ── Startup (local dev only: `python app.py`) ──────────────────────────────────
if __name__ == '__main__':
    with app.app_context():
        added  = seed_db(db, Camera)          # seeding is LOCAL ONLY — never runs on the server
        filled = backfill_costs(db, Camera)
        if added:
            print(f'[seed] Added {added} cameras to the database.')
        if filled:
            print(f'[seed] Backfilled import cost for {filled} sale cameras.')
    app.run(host='0.0.0.0', port=5000, debug=False)
