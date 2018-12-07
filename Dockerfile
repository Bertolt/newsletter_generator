# Use an official Python runtime as a parent image
FROM ubuntu

# Set the working directory to /app
WORKDIR /app

# Copy the current directory contents into the container at /app
ADD . /app

# Install any needed packages specified in requirements.txt


# Make port 80 available to the world outside this container


# Define environment variable


# Run app.py when the container launches
