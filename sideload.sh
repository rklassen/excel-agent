#!/bin/zsh

# Sideload Office.js Add-in for Excel on Mac
# This script clears the add-in cache and copies the manifest to the sideload folder

# Get the current user's home directory
USER_HOME="$HOME"

# Define the sideload folder path
SIDELOAD_DIR="$USER_HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

# Define the manifest file path (assuming script is run from project root)
MANIFEST_FILE="./manifest.xml"

# Check if manifest.xml exists
if [[ ! -f "$MANIFEST_FILE" ]]; then
    echo "Error: manifest.xml not found in current directory."
    echo "Please run this script from the excel-agent project root."
    exit 1
fi

# Create the sideload directory if it doesn't exist
if [[ ! -d "$SIDELOAD_DIR" ]]; then
    echo "Creating sideload directory: $SIDELOAD_DIR"
    mkdir -p "$SIDELOAD_DIR"
    if [[ $? -ne 0 ]]; then
        echo "Error: Failed to create sideload directory."
        exit 1
    fi
fi

# Clear the add-in cache (delete contents of wef folder)
echo "Clearing add-in cache..."
rm -rf "$SIDELOAD_DIR"/*
if [[ $? -ne 0 ]]; then
    echo "Warning: Failed to clear cache, but continuing..."
fi

# Copy the manifest to the sideload folder
echo "Copying manifest.xml to sideload folder..."
cp "$MANIFEST_FILE" "$SIDELOAD_DIR/"
if [[ $? -ne 0 ]]; then
    echo "Error: Failed to copy manifest.xml to $SIDELOAD_DIR"
    exit 1
fi

echo "Manifest sideloaded successfully!"
echo "Location: $SIDELOAD_DIR/manifest.xml"
echo ""
echo "Next steps:"
echo "1. Restart Excel completely (quit and reopen)"
echo "2. The add-in should appear under Developer > Add-ins"
echo ""
echo "Note: Ensure you're signed in with a Microsoft 365 business/education account"
echo "Personal accounts may still block sideloading."