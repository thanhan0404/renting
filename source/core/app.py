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
        # Retrieve form data
        camera_id = request.form.get('camera_id')
        customer_name = request.form.get('customer_name')
        start_date_str = request.form.get('start_date')
        end_date_str = request.form.get('end_date')

        # Convert date strings to datetime objects
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')

        # Calculate rental duration (minimum 1 day)
        days = (end_date - start_date).days
        if days < 1:
            days = 1
            
        # Calculate total price
        camera = Camera.query.get(camera_id)
        total_price = days * camera.price

        # Save to database
        booking = RentalBooking(
            camera_id=camera_id,
            customer_name=customer_name,
            start_date=start_date,
            end_date=end_date,
            total_price=total_price
        )
        db.session.add(booking)
        db.session.commit()
        
        message = f"Success! You booked the {camera.name} for {days} day(s). Total: ${total_price}"
    
    # Fetch cameras meant for rent
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
            
            # Create Camera objects
            cam1 = Camera(name="Sony A7III", type="Sale", price=1999.99, stock=5)
            cam2 = Camera(name="Canon EOS R5", type="Rent", price=85.00, stock=2)
            cam3 = Camera(name="Nikon Z7 II", type="Rent", price=70.00, stock=3)
            cam4 = Camera(name="Fujifilm X-T4", type="Sale", price=1699.00, stock=0) # Out of stock example
            
            # Add them to the session
            db.session.add_all([cam1, cam2, cam3, cam4])
            
            # Commit the session to save them to the SQLite file
            db.session.commit()
            print("Dummy cameras added successfully!")
        else:
            print("Database already contains data. Skipping dummy data insertion.")

    # Start the server
    app.run(host='0.0.0.0', port=5000, debug=False)