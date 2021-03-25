from __main__ import *
import tkinter as tk # GUI
from tkinter.filedialog import asksaveasfile # GUI
import pyqrcode # qr code generation
import png # qr code picture saving
import labels # pip install pylabels
import os
import sys
from reportlab.graphics import shapes # for pylabel labels
from reportlab.graphics.shapes import Image # for pylabel labels
import random, string


# Set Template Options. Add needed options
template_options = ['Avery_6467', 'Avery_94214', 'Avery_5161']
fiber_options = ["C3", "C4", "C5", "D3", "D4", "D5"]  # list entries for fiber options dropdown
gsm_options = ["69", "76", "120", "125"]  # list entries for gsm options dropdown
manyear_options = ["P20", "P21", "A21", "F21"] # list entries for manufacture year

# Generate URL for QRcode
def generate_url(part_num):
  url = 'https://www.physics.purdue.edu/cmsfpix/CMSTrackerFlatSheetsDB_Data/'
  url += str(part_num) + '/' + str(part_num) + '.pdf'
  return url

# GUI and rest of program runs here
class Flatplatedata(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.initialize_user_interface()

    # Function initializes the user interface (GUI)
    def initialize_user_interface(self):
        self.parent.geometry("500x500")
        self.parent.title("Flatplate label/QR generator")

        # Title Label
        lab1 = tk.Label(self.parent, text="Enter Flatplate Info", fg='purple', bg="yellow", relief='solid',
                        font=("Helvetica", "14", "bold"))
        lab1.place(x=170, y=0)

        # Allows User to choose template type
        self.template_lab = tk.Label(self.parent, text='Print Template Type:', fg='black', relief='sunken',
                                     font=('Helvetica', '12', 'bold'))
        self.template_lab.place(x=70, y = 35)
        self.template_str = tk.StringVar(self.parent) # contains value
        self.template_str.set(template_options[0])
        self.template_type = tk.OptionMenu(self.parent, self.template_str, *template_options)
        self.template_type.place(x=300, y=38)

        # Label and Field for Fiber Type
        self.fiblab = tk.Label(self.parent, text="Fiber type & ply #:", fg='black', relief="sunken",
                               font=("Helvetica", "12", "bold"))
        self.fiblab.place(x=70, y=80)
        self.fibtype_str = tk.StringVar(self.parent) # contains value
        self.fibtype_str.set(fiber_options[0])
        self.fibtype = tk.OptionMenu(self.parent, self.fibtype_str, *fiber_options)
        self.fibtype.place(x=300, y=83)

        # Label and Field GSM
        self.gsmlab=tk.Label(self.parent,text="GSM:",fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.gsmlab.place(x=70, y=125)
        self.gsmtype_str = tk.StringVar(self.parent) # contains value
        self.gsmtype_str.set(gsm_options[0])
        self.gsmtype = tk.OptionMenu(self.parent, self.gsmtype_str, *gsm_options)
        self.gsmtype.place(x=300, y=128)

        # Label and Field Lot
        self.lotlab = tk.Label(self.parent, text="Lot #:", fg='black', relief="sunken",
                              font=("Helvetica", "12", "bold"))
        self.lotlab.place(x=70, y=170)
        self.lottype=tk.Entry(self.parent) # contains value
        self.lottype.insert(0, "#####-#")
        self.lottype.place(x=300, y=173)

        # Label and Field for Manufacture Year
        self.manyearlab = tk.Label(self.parent, text="Manufacturer & Year:", fg='black', relief="sunken",
                                   font=("Helvetica", "12", "bold"))
        self.manyearlab.place(x=70, y=215)
        self.manyearlab_str = tk.StringVar(self.parent) # contains value
        self.manyearlab_str.set(manyear_options[0])
        self.manyeartype = tk.OptionMenu(self.parent, self.manyearlab_str, *manyear_options)
        self.manyeartype.place(x=300, y=218)

        # Label and Field for Part Identifier
        self.quantlab = tk.Label(self.parent,  text="Range of plates (start, end):", fg='black', relief="sunken",
                                 font=("Helvetica", "12", "bold"))
        self.quantlab.place(x=70, y=260)
        # Minimum quant field
        self.quant_low = tk.Entry(self.parent) # contains value
        self.quant_low.insert(0, "1")
        self.quant_low.place(x=300, y=260, width=50)
        # Maximum quant field
        self.quant_high = tk.Entry(self.parent) # contains value
        self.quant_high.insert(0, "1")
        self.quant_high.place(x=355, y=260, width=50)

        # Generate Button
        global prevButton
        # Calls the function generate_codes()
        prevButton = tk.Button(self.parent, text="Generate Labels", fg='black', bg='green',
                               relief="ridge", font=("Helvetica", "14", "bold"), command=self.generate_codes)
        prevButton.place(x=175, y=435)


    # Takes input from GUI to prepare for printing
    def generate_codes(self):
      # Ask for path
      # QR codes are generated in this path, which are read by the program to generate a PDF containing all of the
      # codes. The program will then automatically delete all of the QR code files.
      self.path = tk.filedialog.askdirectory(title='Select Folder')
      print(self.path)

      quant_lo = 1
      quant_lo = int(self.quant_low.get())
      quant_hi = 1
      quant_hi = int(self.quant_high.get())

      # Generate product code and QRcode
      codeset = []  # list contains generated product codes
      qrset = []    # list contains path to qrcode images
      code = ''

      for i in range(quant_lo, quant_hi + 1):
        unique_code = random.choice(string.ascii_letters) # throws in a unique code at the end for qrcodes
        # Get Product Code
        code = "{}-{}-{}-{}-{}".format(self.fibtype_str.get(), self.gsmtype_str.get(), self.lottype.get(), self.manyearlab_str.get(), i)
        codeset.append(code)
        # Generate QRcode
        url = generate_url(code)
        qr = pyqrcode.create(url)
        #qr_path = self.path + '/' + str(code) + unique_code + '.png'  # for running on macOS
        qr_path = self.path + '\\' + str(code) + '.png' # for running on Windows
        qr.png(qr_path, scale=2)
        qrset.append(qr_path)

      # Print Preview on GUI
      preview = tk.Label(self.parent, text="Preview:", fg='black', font=("Helvetica", "10", "bold"))  # preview label
      preview.place(x=75, y=340)
      codePrev = tk.Label(self.parent, text=code, fg='black', relief="sunken", font=("Helvetica", "24", "bold")) #product code label
      codePrev.place(x=70, y=370)

      # Call get_pdf() function to send in codeset and qrset
      self.get_pdf(codeset, qrset)


    # Generate the pdf
    def get_pdf(self, codeset, qrset):
      # Get custom dimensions depending on the tempalte
      (label_h, label_w, top_mar, bot_mar, left_mar, right_mar, num_col, num_row, cgap, rgap) = self.get_template_dim()
      # Set up the specs for Letter Paper and Tempalte Type
      specs = labels.Specification(215.9, 279.4, num_col, num_row, label_w, label_h, corner_radius=1,
              top_margin=top_mar, bottom_margin=bot_mar, left_margin=left_mar, right_margin=right_mar, column_gap=cgap, row_gap=rgap)

      # Initialize the sheet that will be used for adding labels
      sheet = labels.Sheet(specs, self.draw_label, border=False) # set border=True for testing with borders. (Turn it off for printing to paper)

      for i in range(len(codeset)):
        # Calls in a dictionary to add both the product code and QR code onto the same label. The dictionary is read
        # in the function draw_label.
        _print_dict = {
                'text' : codeset[i],
                'image' : qrset[i],
            }
        sheet.add_label(_print_dict)

      filename = "{}-{}-{}-{}".format(self.fibtype_str.get(), self.gsmtype_str.get(), self.lottype.get(), self.manyearlab_str.get())
      #sheet.save(self.path + '/' + filename + '_' + self.template_str.get() + '.pdf')  # macs
      sheet.save(self.path + '\\' + filename + '_' +self.template_str.get() + '.pdf') # windows

      # Delete files in path
      for i in qrset:
        os.remove(i)

      print('Print Complete')


    # Draw two objects into label: product code and QRcode
    # The size and font changes depending on the template
    # shapes.String -> the produce code saved in _print_dict
    # Image() -> the QRcode png accessed through filepath stored under 'image' in _print_dict
    def draw_label(self, label, width, height, obj):
      # 4 * 20
      if self.template_str.get() == 'Avery_6467':
        txt = obj.get('text')
        if len(txt) > 22:
          t1 = txt[0:len(txt)//2]
          t2 = txt[len(txt)//2 if len(txt)%2 == 0
                                 else ((len(txt)//2)+1):]
          label.add(shapes.String(7, 20, t1, fontName="Helvetica", fontSize=7))
          label.add(shapes.String(7, 5, t2, fontName="Helvetica", fontSize=7))
        else:
          label.add(shapes.String(5, 15, obj.get('text'), fontName="Helvetica", fontSize=7))
        label.add(Image(85, 5, 29, 29, obj.get('image')))
      # 2 * 10
      elif self.template_str.get() == 'Avery_5161':
        label.add(shapes.String(20, 30, obj.get('text'), fontName="Helvetica", fontSize=13))
        label.add(Image(180, 12, 55, 55, obj.get('image')))
      # 2 * 16
      elif self.template_str.get() == 'Avery_94214':
        label.add(shapes.String(10, 15, obj.get('text'), fontName="Helvetica", fontSize=10))
        label.add(Image(160, 2, 40, 40, obj.get('image')))

    # return template dim -> (label_h, label_w, top_mar, bot_mar, left_mar, right_mar, num_col, num_row, col_gap, row_gap)
    # All measurements are in mm
    def get_template_dim(self):
      template = self.template_str.get()
      if template == 'Avery_6467':
        return (12.7, 44.5, 12.7, 12.7, 9.6, 7, 4, 20, 7.1, 0)
      elif template == 'Avery_94214':
        return (15.875, 76.2, 12.7, 12.7, 21, 21.49, 2, 16, 21.01, 0)
      elif template == 'Avery_5161':
        return (25.4, 101.8, 12.7, 12.7, 5, 3.3, 2, 10, 4, 0)


if __name__ == '__main__':
  root = tk.Tk()
  run = Flatplatedata(root)
  root.mainloop()