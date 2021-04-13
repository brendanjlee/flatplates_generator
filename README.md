# flatplates_generator

Create QR Code labels for Flatplates with Python.

Uses a simple GUI to produce sheets of QR Code labels on following formats: Avery 6467, Avery 5161, and Avery 94214.<br>

##### Table of Contents
* [Technologies and Dependencies](#technologies-and-dependencies)
* [Running the Code](#running-the-code)
* [GUI Demo](#demo-of-gui)
* [Code Examples](#code-examples)
* [Basic Program Structure](#Basic-Flow-of-Program)


## Technologies and Dependencies
### Dependencies:<br>
You must import these packages using PIP or Conda in order to modify and run the program.
```
in qrGenerator.py:
```
```
import tkinter as tk
import pyqrcode
import png
import labels
from reportlab.graphics import shapes
from reportlab.graphics.shapes import Image
```
### Creating an executable using PyInstaller
This project was created with the help of [this](https://jacob-brown.github.io/2019-09-10-pyinstaller/) tutorial.
#### Installing PyInstaller
```
pip install pyinstaller
```
#### Creating an Executable for Windows
```
$ pyinstaller --windowed --onefile qrGenerator.py
```
#### Creating an Executable for Mac
```
pyinstaller --windowed  DogsGo.py
```

----------------------
## Running the Code
If you want to run the code without creating an executable, you can simpy type
```
python qrGenerator.py
```
in the terminal. This should prompt a window.

----------------------
## Demo of GUI
<br>
<img width="507" alt="Screen Shot 2021-03-16 at 4 45 00 PM" src="https://user-images.githubusercontent.com/31088155/111377377-f3d4bb00-8676-11eb-91f1-1e3439fe6cef.png">

----------------------
## Basic Flow of Program
* the main function calls tk() which sets up the GUI
* User input is taken from ```initialize_user_interface()``` which calls ```generate_codes()```
* ```generate_codes()``` will then generate the qrLabels in a loop and then save them to the temp. files. These qr code files will be automatically deleted when the PDF is generated. The function ```get_pdf()``` will be called
* ```get_pdf()``` will use pylabels to draw the pdf using ```draw_label``` and ```get_template_dim()```

----------------------
## Code Examples
```qrGenerator.py```
### Template Options
To set options for template, fiber, gsm, and manyufacture year, add to these lists in :
![image](https://user-images.githubusercontent.com/31088155/112562359-76632600-8dad-11eb-80d5-202f7f430cca.png)

To change the url for the QR codes, modify the function ```denerate_url()``` :
![image](https://user-images.githubusercontent.com/31088155/112562556-e4a7e880-8dad-11eb-9647-048178bebf1e.png)

### Per System Output Path
While you can select where to save to when running the executable, you have to switch one line of code depending on the system that you are running. The line to change can be found at ```generate_codes()``` function.
<img width="927" alt="Screen Shot 2021-04-13 at 2 32 05 PM" src="https://user-images.githubusercontent.com/31088155/114602806-084fa780-9c65-11eb-8bac-3993d1b22fe4.png">

### Adding Different Templates
The different label dimensions can be found in the ```get_template_dim()``` function where the dimensions were manually inputted. To ensure that the newly inputted dimensions work, you also have to make some margin changes to
*  ```draw_label``` which actually draws the labels with the given input

I would recommend taking a look at [this](https://github.com/bcbnz/pylabels/blob/master/labels/specifications.py) which will explain every parameter that pylabels will take.
