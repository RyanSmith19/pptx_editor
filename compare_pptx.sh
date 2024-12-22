#!/bin/bash

# Ensure correct number of arguments
if [ "$#" -ne 2 ]; then
    echo "Usage: $0 <file_or_directory1> <file_or_directory2>"
    exit 1
fi

# Input paths
INPUT1="$1"
INPUT2="$2"

# Temporary directories for extraction
TEMP_DIR1="temp_pptx1"
TEMP_DIR2="temp_pptx2"

# Create temporary directories
mkdir -p "$TEMP_DIR1"
mkdir -p "$TEMP_DIR2"

# Function to compare two PPTX files
compare_pptx_files() {
    local file1="$1"
    local file2="$2"
    local temp_dir1="$TEMP_DIR1/$(basename "$file1" .pptx)"
    local temp_dir2="$TEMP_DIR2/$(basename "$file2" .pptx)"
    
    # Create temporary subdirectories
    mkdir -p "$temp_dir1"
    mkdir -p "$temp_dir2"

    # Extract contents of PPTX files
    unzip -q "$file1" -d "$temp_dir1"
    unzip -q "$file2" -d "$temp_dir2"

    # Compare XML content and show meaningful differences
    echo "Comparing: $(basename "$file1")"
    diff -r "$temp_dir1" "$temp_dir2" \
        | grep -E '^<|^>' \
        | grep -v '<?xml version' \
        | grep -v 'standalone="yes"' \
        | grep -v '^<Types xmlns' \
        | head -n 20
    echo "---"

    # Clean up temporary subdirectories
    rm -rf "$temp_dir1"
    rm -rf "$temp_dir2"
}

# If both inputs are directories, compare matching files
if [ -d "$INPUT1" ] && [ -d "$INPUT2" ]; then
    echo "Comparing files in directories: $INPUT1 and $INPUT2"

    # Iterate through .pptx files in the first directory
    for file1 in "$INPUT1"/*.pptx; do
        file2="$INPUT2/$(basename "$file1")"
        
        # Check if corresponding file exists in the second directory
        if [ -f "$file2" ]; then
            compare_pptx_files "$file1" "$file2"
        else
            echo "No matching file for $(basename "$file1") in $INPUT2"
        fi
    done
# If inputs are files, compare them directly
elif [ -f "$INPUT1" ] && [ -f "$INPUT2" ]; then
    compare_pptx_files "$INPUT1" "$INPUT2"
else
    echo "Error: Both inputs must be either files or directories."
fi

# Clean up temporary directories
rm -rf "$TEMP_DIR1"
rm -rf "$TEMP_DIR2"

echo "Comparison complete."

