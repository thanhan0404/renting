from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Camera(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    type = db.Column(db.String(20)) # 'Sale' or 'Rent'
    price = db.Column(db.Float, nullable=False) # Purchase price or Price per day
    stock = db.Column(db.Integer, default=1)

class RentalBooking(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    camera_id = db.Column(db.Integer, db.ForeignKey('camera.id'))
    customer_name = db.Column(db.String(100))
    start_date = db.Column(db.DateTime, nullable=False)
    end_date = db.Column(db.DateTime, nullable=False)
    total_price = db.Column(db.Float)