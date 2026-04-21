import os
from flask_admin import Admin
from flask_admin.contrib.sqla import ModelView
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for
from models import db, Camera, RentalBooking

basedir = os.path.abspath(os.path.dirname(__file__))
template_dir = os.path.join(basedir, '..', 'templates')
app = Flask(__name__, template_folder=template_dir)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///camerashop.db'
app.config['SECRET_KEY'] = 'a_secure_secret_key' # Flask-Admin requires a secret key for session security
db.init_app(app)

admin = Admin(app, name='Camera Shop Admin')

# Add administrative views for your models
admin.add_view(ModelView(Camera, db.session))
admin.add_view(ModelView(RentalBooking, db.session))

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/store')
def store():
    # Fetch only cameras meant for sale
    sale_cameras = Camera.query.filter_by(type='Sale').all()
    return render_template('store.html', cameras=sale_cameras)

@app.route('/rent', methods=['GET', 'POST'])
def rent():
    message = None
    if request.method == 'POST':
        camera_id = request.form.get('camera_id')
        customer_name = request.form.get('customer_name')
        phone = request.form.get('phone') # Capture phone
        notes = request.form.get('notes') # Capture notes
        start_date_str = request.form.get('start_date')
        end_date_str = request.form.get('end_date')

        # Notice the 'T' and %H:%M added to the strptime format
        start_date = datetime.strptime(start_date_str, '%Y-%m-%dT%H:%M')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%dT%H:%M')

        # Calculate rental duration in hours, then convert to days (minimum 1 day)
        duration_seconds = (end_date - start_date).total_seconds()
        days = duration_seconds / 86400.0 # 86400 seconds in a day
        if days < 1:
            days = 1
            
        camera = Camera.query.get(camera_id)
        # Round the price to 2 decimal places
        total_price = round(days * camera.price, 2)

        booking = RentalBooking(
            camera_id=camera_id,
            customer_name=customer_name,
            phone=phone, # Save to DB
            notes=notes, # Save to DB
            start_date=start_date,
            end_date=end_date,
            total_price=total_price
        )
        db.session.add(booking)
        db.session.commit()
        
        message = f"Thành công! Bạn đã đặt thuê {camera.name}. Tổng tiền dự kiến: ${total_price}"
    
    rent_cameras = Camera.query.filter_by(type='Rent').all()
    return render_template('rent.html', cameras=rent_cameras, message=message)

if __name__ == '__main__':
    # 1. Open the application context
    with app.app_context():
        # 2. Create the tables based on models.py
        db.create_all() 
        
        # 3. Check if the database is empty before adding dummy data
        if not Camera.query.first():
            print("Database is empty. Adding dummy cameras...")
            cam1 = Camera(name="Sony A7III", type="Sale", price=1999.99, stock=5)
            cam2 = Camera(name="Canon EOS R5", type="Rent", price=85.00, stock=2)
            cam3 = Camera(name="Nikon Z7 II", type="Rent", price=70.00, stock=3)
            cam4 = Camera(name="Fujifilm X-T4", type="Sale", price=1699.00, stock=0) # Out of stock example
            
            # Add them to the session
            db.session.add_all([cam1, cam2, cam3, cam4])
            db.session.commit()
            print("Dummy cameras added successfully!")
        else:
            print("Database already contains data. Skipping dummy data insertion.")

    # Start the server
    app.run(host='0.0.0.0', port=5000, debug=False)