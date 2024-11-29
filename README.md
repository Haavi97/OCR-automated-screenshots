# OCR Scanner

A desktop application that automates the process of taking screenshots of documents, performing OCR (Optical Character Recognition), and saving the results as an EPUB file.

## Features

- Screenshot capture with region selection
- Automated page navigation
- OCR text extraction
- EPUB file generation with formatted text
- Pause/Resume functionality
- Progress tracking
- Custom save location
- User-friendly interface

## Prerequisites

- Python 3.8 or higher
- Tesseract OCR engine
- Windows OS (currently optimized for Windows)

## Installation

1. Install Tesseract OCR for Windows from the [official repository](https://github.com/UB-Mannheim/tesseract/wiki)

2. Clone this repository:
```bash
git clone https://github.com/Haavi97/OCR-automated-screenshots.git
cd ocr-automated-screenshots
```

3. Create and activate a virtual environment:
```bash
python -m venv venv
venv\Scripts\activate
```

4. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python main.py
```

2. Set up your scan:
   - Choose the number of pages
   - Set delay between pages
   - Select save location for the EPUB file
   - Select the region to scan
   - Select the "Next" button location

3. Click "Start" to begin the scanning process

4. Use the "Pause" button if needed during scanning

## Building the Executable

To create a standalone executable:

```bash
pyinstaller --name="OCR Scanner" --windowed --onefile main.py
```

The executable will be created in the `dist` folder.

Alternatively you can use the scripts `build_unix.sh` and `build_windows.sh`

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) for OCR functionality
- [EbookLib](https://github.com/aerkalov/ebooklib) for EPUB generation
