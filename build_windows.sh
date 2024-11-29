# First, clear any old builds
rm -rf build dist

pyinstaller --name="OCR Scanner" --windowed --onefile screenshot-ocr-automation.py