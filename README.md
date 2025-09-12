# TaxApp
python3 MyApp.py

python3 -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip
pip install --pre pyinstaller
pip install -r requirements.txt

pyinstaller --name "TaxApp" --onefile --windowed --icon icon.ico MyApp.py
