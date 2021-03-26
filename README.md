# flatplates_generator

Create QR Code labels for Flatplates with Python.

Uses a simple GUI to produce sheets of QR Code labels on following formats: Avery 6467, Avery 5161, and Avery 94214.<br>

Dependencies:<br>
```
import tkinter as tk
import pyqrcode
import png
import labels
from reportlab.graphics import shapes
from reportlab.graphics.shapes import Image
```

Demo of GUI
<br>
<img width="507" alt="Screen Shot 2021-03-16 at 4 45 00 PM" src="https://user-images.githubusercontent.com/31088155/111377377-f3d4bb00-8676-11eb-91f1-1e3439fe6cef.png">

---
## How to Modify the Code
```qrGenerator.py```
### Template Options
To set options for template, fiber, gsm, and manyufacture year, add to these lists in :
![image](https://user-images.githubusercontent.com/31088155/112562359-76632600-8dad-11eb-80d5-202f7f430cca.png)

To change the url for the QR codes, modify the function ```denerate_url()``` :
![image](https://user-images.githubusercontent.com/31088155/112562556-e4a7e880-8dad-11eb-9647-048178bebf1e.png)
