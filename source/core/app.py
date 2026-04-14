import os
from flask import Flask, render_template, request
from models import db, Camera

basedir = os.path.abspath(os.path.dirname(__file__))
template_dir = os.path.join(basedir, '..', 'templates')

app = Flask(__name__, template_folder=template_dir)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///camerashop.db'
db.init_app(app)
print("FLASK IS LOOKING FOR TEMPLATES HERE:", template_dir)
# -----------------------------------------------------------------------------
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
    if request.method == 'POST':
        # Logic to handle the booking form submission goes here
        pass
    
    # Fetch cameras meant for rent
    rent_cameras = Camera.query.filter_by(type='Rent').all()
    return render_template('rent.html', cameras=rent_cameras)

if __name__ == '__main__':
    with app.app_context():
        db.create_all() # Creates the database file
    app.run(host='0.0.0.0', port=5000, debug=False)
