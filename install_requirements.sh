#!/bin/bash

# Check if requirements.txt exists
if [ ! -f requirements.txt ]; then
  echo "Error: requirements.txt not found in the current directory."
  exit 1
fi

# Install the required libraries
echo "Installing required libraries..."
pip install -r requirements.txt

# Check if the installation was successful
if [ $? -eq 0 ]; then
  echo "All libraries installed successfully!"
else
  echo "Error: Failed to install one or more libraries."
  exit 1
fi

