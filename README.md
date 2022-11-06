# qrexcel-generator

Qrexcel is a Python script that generates a QR Code for every user having his provided info in the Excel file like name, phone, type, and ID. and adds a specific logo, uploads it to a particular folder on Mega.nz, and finally adds the link of the uploaded QR Image to the original Excel file and save a local copy of the QR Code Image in a new folder.

## Installation

#### After cloning the repo, Use the package manager [pip](https://pip.pypa.io/en/stable/) to install these required libraries:

```bash
pip install openpyxl
pip install qrcode
pip install PIL
pip install mega.py
```

## Usage

Make sure that your folder has the script and the Excel file named 'data.xlsx' and your logo named 'logo.jpg' all without quotes. if you want to pass another name you can change it here:
```python
Line 11: logo = Image.open('logo.jpg')
Line 12: file = openpyxl.load_workbook('data.xlsx')
Line 69: file.save('../data.xlsx')
```
 Here you can change these mega.nz credentials with yours:

```python
Line 36: m = mega.login("sitew99844@hempyl.com", "TheSant0x")
```
Here you can choose the mega folder's name which will be created by the script:
``` python
Line 37: folderName='Generated QR Codes'
```
##### Note that the daily uploading limit for the free plan is 5GB.
