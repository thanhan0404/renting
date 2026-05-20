import os
from flask_admin import Admin
from flask_admin.contrib.sqla import ModelView
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, jsonify

# Import thêm PurchaseOrder
from models import db, Camera, RentalBooking, PurchaseOrder

basedir = os.path.abspath(os.path.dirname(__file__))
template_dir = os.path.join(basedir, '..', 'templates')
app = Flask(__name__, template_folder=template_dir)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///camerashop.db'
app.config['SECRET_KEY'] = 'a_secure_secret_key' 
db.init_app(app)

admin = Admin(app, name='Camera Shop Admin')

# Thêm views cho Admin
admin.add_view(ModelView(Camera, db.session))
admin.add_view(ModelView(RentalBooking, db.session))
admin.add_view(ModelView(PurchaseOrder, db.session)) # View cho đơn mua hàng

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/store')
def store():
    sale_cameras = Camera.query.filter_by(type='Sale').all()
    return render_template('store.html', cameras=sale_cameras)

# --- CHỨC NĂNG GIỎ HÀNG VÀ THANH TOÁN ---

@app.route('/add_to_cart/<int:camera_id>')
def add_to_cart(camera_id):
    if 'cart' not in session:
        session['cart'] = {}
    
    cart = session['cart']
    camera = Camera.query.get_or_404(camera_id)
    
    if camera.stock > 0:
        cam_id_str = str(camera_id)
        if cam_id_str in cart:
            if cart[cam_id_str] < camera.stock:
                cart[cam_id_str] += 1
        else:
            cart[cam_id_str] = 1
        session.modified = True
        
    return redirect(url_for('cart'))

@app.route('/cart')
def cart():
    cart_items = session.get('cart', {})
    items = []
    total = 0
    for cam_id, qty in cart_items.items():
        camera = Camera.query.get(int(cam_id))
        if camera:
            item_total = camera.price * qty
            total += item_total
            items.append({'camera': camera, 'quantity': qty, 'item_total': item_total})
    return render_template('cart.html', items=items, total=total)

@app.route('/checkout', methods=['POST'])
def checkout():
    cart_items = session.get('cart', {})
    if not cart_items:
        return redirect(url_for('store'))
    
    customer_name = request.form.get('customer_name')
    phone = request.form.get('phone')
    address = request.form.get('address')
    
    for cam_id, qty in cart_items.items():
        camera = Camera.query.get(int(cam_id))
        if camera and camera.stock >= qty:
            # Trừ số lượng tồn kho
            camera.stock -= qty
            
            # Tạo đơn hàng mới
            order = PurchaseOrder(
                camera_id=camera.id,
                customer_name=customer_name,
                phone=phone,
                address=address,
                quantity=qty,
                total_price=camera.price * qty
            )
            db.session.add(order)
            
    db.session.commit()
    session.pop('cart', None) # Xóa giỏ hàng sau khi thanh toán
    return render_template('cart.html', message="Mua hàng thành công! Chúng tôi sẽ sớm liên hệ với bạn.", items=[], total=0)

# --- CHỨC NĂNG THUÊ MÁY (Giữ nguyên) ---

@app.route('/api/camera-bookings/<int:camera_id>')
def camera_bookings_api(camera_id):
    """Return booked date ranges for a camera so the frontend can display unavailable periods."""
    bookings = RentalBooking.query.filter_by(camera_id=camera_id).all()
    return jsonify([{
        'start': b.start_date.isoformat(),
        'end': b.end_date.isoformat()
    } for b in bookings])


@app.route('/api/available-cameras')
def available_cameras_api():
    """Return cameras with no overlapping booking for the requested date range."""
    start_str = request.args.get('start')
    end_str = request.args.get('end')

    if not start_str or not end_str:
        return jsonify({'error': 'Missing start or end date'}), 400

    try:
        start = datetime.strptime(start_str, '%Y-%m-%dT%H:%M')
        end = datetime.strptime(end_str, '%Y-%m-%dT%H:%M')
    except ValueError:
        return jsonify({'error': 'Invalid date format'}), 400

    if end <= start:
        return jsonify({'error': 'End date must be after start date'}), 400

    # Camera IDs that already have a booking overlapping [start, end)
    busy_ids = [r[0] for r in db.session.query(RentalBooking.camera_id).filter(
        RentalBooking.start_date < end,
        RentalBooking.end_date > start
    ).distinct().all()]

    if busy_ids:
        cameras = Camera.query.filter_by(type='Rent').filter(~Camera.id.in_(busy_ids)).all()
    else:
        cameras = Camera.query.filter_by(type='Rent').all()

    return jsonify([{'id': c.id, 'name': c.name, 'price': c.price} for c in cameras])


@app.route('/rent', methods=['GET', 'POST'])
def rent():
    message = None
    if request.method == 'POST':
        camera_id = request.form.get('camera_id')
        customer_name = request.form.get('customer_name')
        phone = request.form.get('phone') 
        notes = request.form.get('notes') 
        start_date_str = request.form.get('start_date')
        end_date_str = request.form.get('end_date')

        start_date = datetime.strptime(start_date_str, '%Y-%m-%dT%H:%M')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%dT%H:%M')

        duration_seconds = (end_date - start_date).total_seconds()
        days = duration_seconds / 86400.0 
        if days < 1:
            days = 1
            
        camera = Camera.query.get(camera_id)
        total_price = round(days * camera.price, 2)

        booking = RentalBooking(
            camera_id=camera_id,
            customer_name=customer_name,
            phone=phone, 
            notes=notes, 
            start_date=start_date,
            end_date=end_date,
            total_price=total_price
        )
        db.session.add(booking)
        db.session.commit()
        
        message = f"Thành công! Bạn đã đặt thuê {camera.name}. Tổng tiền dự kiến: ${total_price}"
    
    rent_cameras = Camera.query.filter_by(type='Rent').all()
    active_flow = request.args.get('flow', 'camera')
    return render_template('rent.html', cameras=rent_cameras, message=message, active_flow=active_flow)

if __name__ == '__main__':
    with app.app_context():
        db.create_all() 
        if not Camera.query.first():
            cam1 = Camera(name="Sony A7III", type="Sale", price=1999.99, stock=5)
            cam2 = Camera(name="Canon EOS R5", type="Rent", price=85.00, stock=2)
            cam3 = Camera(name="Nikon Z7 II", type="Rent", price=70.00, stock=3)
            cam4 = Camera(name="Fujifilm X-T4", type="Sale", price=1699.00, stock=0) 
            db.session.add_all([cam1, cam2, cam3, cam4])
            db.session.commit()
    app.run(host='0.0.0.0', port=5000, debug=False)