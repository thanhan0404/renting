import os
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for
from models import db, Camera, RentalBooking

basedir = os.path.abspath(os.path.dirname(__file__))
template_dir = os.path.join(basedir, '..', 'templates')

app = Flask(__name__, template_folder=template_dir)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///camerashop.db'
db.init_app(app)
print("FLASK IS LOOKING FOR TEMPLATES HERE:", template_dir)

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
    with app.app_context():
        db.create_all()
        # Add dummy data if the database is empty
        if not Camera.query.first():
            db.session.add(Camera(name="Sony A7III", type="Sale", price=1999.99, stock=5))
            db.session.add(Camera(name="Canon EOS R5", type="Rent", price=85.00, stock=2))
            db.session.add(Camera(name="Nikon Z7 II", type="Rent", price=70.00, stock=3))
            db.session.commit()
            print("Dummy cameras added to database!")
            
    app.run(host='0.0.0.0', port=5000, debug=False)