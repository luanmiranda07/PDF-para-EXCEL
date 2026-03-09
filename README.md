uv run pyinstaller --onefile --hidden-import='pdfminer' --hidden-import='PIL' main.py

pyinstaller --onefile --hidden-import='pdfminer' --hidden-import='PIL' main.py

pip install --upgrade pyinstaller pdfplumber

o que funciona perfeitamente

pyinstaller --clean --noconfirm --onefile --windowed --icon=favicon.ico --hidden-import pdfminer --hidden-import PIL main.py
