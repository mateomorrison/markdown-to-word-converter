import os
from pathlib import Path
import pypandoc

def convert_files():
    # Ensure pandoc is installed
    try:
        # This downloads pandoc if not found on the system
        pypandoc.ensure_pandoc_installed()
    except OSError:
        print("Could not find or install pandoc automatically.")
        print("Please install it manually or check your internet connection.")
        return

    # Define directories
    input_dir = Path("input")
    output_dir = Path("output")

    # Create output directory if it doesn't exist
    output_dir.mkdir(exist_ok=True)

    # Iterate over all markdown files in the input directory
    for file_path in input_dir.glob("*.md"):
        # Define the output filename with .docx extension
        output_filename = file_path.stem + ".docx"
        output_path = output_dir / output_filename
        
        print(f"Converting '{file_path}' to '{output_path}'...")
        
        try:
            # Convert the file using pypandoc
            pypandoc.convert_file(str(file_path), 'docx', outputfile=str(output_path))
            print("Success!")
        except RuntimeError as e:
            print(f"Error converting {file_path}: {e}")

if __name__ == "__main__":
    convert_files()