# Use an official Python runtime as a parent image
FROM python:3.11-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file and install dependencies
COPY backend/requirements.txt ./backend/
RUN pip install --no-cache-dir -r backend/requirements.txt

# Install Gunicorn for production serving
RUN pip install --no-cache-dir gunicorn

# Copy the frontend and backend directories into the container
COPY frontend/ ./frontend/
COPY backend/ ./backend/

# Set the working directory to where app.py is located
WORKDIR /app/backend

# Expose port 5000 to the outside world
EXPOSE 5000

# Run the app using Gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"]