# Use the official Python image from the Docker Hub
FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Create and set the working directory
WORKDIR /app

# Copy the requirements file into the container
COPY requirements.txt /app/

# Install the Python dependencies
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copy the entire application into the container
COPY . /app/

# Expose port 5000 for the Flask application
EXPOSE 5000

# Define the command to run the Flask application
CMD ["flask", "run", "--host=0.0.0.0", "--port=5000"]
