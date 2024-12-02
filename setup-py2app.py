from setuptools import setup

APP = ['your_script_name.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['tkinter', 'PIL', 'pytesseract', 'keyboard', 'ebooklib'],
    'plist': {
        'CFBundleName': "OCR Scanner",
        'CFBundleDisplayName': "OCR Scanner",
        'CFBundleGetInfoString': "Making OCR scanning easy",
        'CFBundleVersion': "1.0.0",
        'CFBundleShortVersionString': "1.0.0",
    }
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
