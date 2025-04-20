import os
import sys
import subprocess
from pathlib import Path

def main():
    # Define paths
    current_dir = Path.cwd()
    py_folder = current_dir / "py"
    xl_folder = current_dir / "xl"

    # Create the output folder if it doesn't exist
    os.makedirs(xl_folder, exist_ok=True)

    # Check if the py folder exists
    if not py_folder.exists():
        print(f"Error: Python scripts folder not found at {py_folder}")
        return 1

    # Get all Python files in the py folder
    python_files = list(py_folder.glob("*.py"))

    if not python_files:
        print(f"No Python files found in {py_folder}")
        return 0

    print(f"Found {len(python_files)} Python scripts to execute")

    # Run each Python file
    for py_file in python_files:
        print(f"\nRunning: {py_file.name}")
        try:
            # Set environment variable for output directory
            env = os.environ.copy()
            env["OUTPUT_DIR"] = str(xl_folder)
            
            # Run script with output directory as argument
            cmd = [sys.executable, str(py_file), "--output", str(xl_folder)]
            result = subprocess.run(cmd, env=env, capture_output=True, text=True)
            
            if result.returncode == 0:
                print(f"✓ Successfully executed {py_file.name}")
                if result.stdout.strip():
                    print(f"Output: {result.stdout.strip()}")
            else:
                print(f"✗ Failed to execute {py_file.name}")
                print(f"Error: {result.stderr.strip()}")
                
        except Exception as e:
            print(f"Error running {py_file.name}: {str(e)}")

    print(f"\nAll scripts executed. Output files should be in {xl_folder}")
    return 0

if __name__ == "__main__":
    sys.exit(main())