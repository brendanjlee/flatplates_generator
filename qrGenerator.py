from __main__ import *
import tkinter as tk
from tkinter.filedialog import asksaveasfile
import pyqrcode
import png
import labels
import os
import sys
from reportlab.graphics import shapes
from reportlab.graphics.shapes import Image

# Set Template Options
template_options = ['Avery_6467', 'Avery_94214', 'Avery_5161']
fiber_options = ["C3", "C4", "C5", "D3", "D4", "D5"]  # list entries for fiber options dropdown
gsm_options = ["69", "76", "120", "125"]  # list entries for gsm options dropdown
manyear_options = ["P20", "P21", "A21", "F21"] # list entries for manufacture year

# Generate URL for QRcode
def generate_url(part_num):
  url = 'https://www.physics.purdue.edu/cmsfpix/CMSTrackerFlatSheetsDB_Data/'
  url += str(part_num) + '/' + str(part_num) + '.pdf'
  return url

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
        lab1 = tk.Label(self.parent, text="Enter Flatplate Info", fg='purple', bg="yellow", relief='solid',font=("Helvetica", "14", "bold"))
        lab1.place(x=170, y=0)

        # Allows User to choose template type
        self.template_lab = tk.Label(self.parent, text='Print Template Type:', fg='black', relief='sunken', font=('Helvetica', '12', 'bold'))
        self.template_lab.place(x=70, y = 35)
        self.template_str = tk.StringVar(self.parent) # contains value
        self.template_str.set(template_options[0])
        self.template_type = tk.OptionMenu(self.parent, self.template_str, *template_options)
        self.template_type.place(x=300, y=38)

        # Label and Field for Fiber Type
        self.fiblab = tk.Label(self.parent, text="Fiber type & ply #:", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
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
        self.lotlab = tk.Label(self.parent, text="Lot #:", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.lotlab.place(x=70, y=170)
        self.lottype=tk.Entry(self.parent) # contains value
        self.lottype.insert(0, "#####-#")
        self.lottype.place(x=300, y=173)

        # Label and Field for Manufacture Year
        self.manyearlab = tk.Label(self.parent, text="Manufacturer & Year:", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.manyearlab.place(x=70, y=215)
        self.manyearlab_str = tk.StringVar(self.parent) # contains value
        self.manyearlab_str.set(manyear_options[0])
        self.manyeartype = tk.OptionMenu(self.parent, self.manyearlab_str, *manyear_options)
        self.manyeartype.place(x=300, y=218)

        # Label and Field for Part Identifier
        self.quantlab = tk.Label(self.parent,  text="Range of plates (start, end):", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
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
        prevButton = tk.Button(self.parent, text="Generate Labels", fg='black', bg='green', relief="ridge", font=("Helvetica", "14", "bold"), command=self.generate_codes)
        prevButton.place(x=175, y=435)

    # Takes input from GUI to prepare for printing
    def generate_codes(self):
      # Ask for path
      self.path = tk.filedialog.askdirectory(title='Select Folder')
      print(self.path)

      quant_lo = 1
      quant_lo = int(self.quant_low.get())
      quant_hi = 1
      quant_hi = int(self.quant_high.get())
      quantity = quant_hi - quant_lo

      # Generate product code and QRcode
      codeset = []  # list contains generated product codes
      qrset = []    # list contains path to qrcode images
      code = ''

      for i in range(quant_lo, quant_hi + 1):
        # Get Product Code
        code = "{}-{}-{}-{}-{}".format(self.fibtype_str.get(), self.gsmtype_str.get(), self.lottype.get(), self.manyearlab_str.get(), i)
        codeset.append(code)
        # Generate QRcode
        url = generate_url(code)
        qr = pyqrcode.create(url)
        # for macs
        #qr_path = self.path + '/' + str(code) + '.png'
        qr_path = self.path + '\\' + str(code) + '.png'
        qr.png(qr_path, scale=2)
        qrset.append(qr_path)

      # Print Preview on GUI
      preview = tk.Label(self.parent, text="Preview:", fg='black', font=("Helvetica", "10", "bold"))  # preview label
      preview.place(x=75, y=340)
      codePrev = tk.Label(self.parent, text=code, fg='black', relief="sunken", font=("Helvetica", "24", "bold")) #product code label
      codePrev.place(x=70, y=370)

      self.get_pdf(codeset, qrset)


    # Generate the pdf
    def get_pdf(self, codeset, qrset):
      (label_h, label_w, top_mar, bot_mar, left_mar, right_mar, num_col, num_row) = self.get_template_dim()

      specs = labels.Specification(215.9, 279.4, num_col, num_row, label_w, label_h, corner_radius=1,
              top_margin=top_mar, bottom_margin=bot_mar, left_margin=left_mar, right_margin=right_mar)
      sheet = labels.Sheet(specs, self.draw_label, border=True)

      for i in range(len(codeset)):
        _print_dict = {
                'text' : codeset[i],
                'image' : qrset[i],
            }
        sheet.add_label(_print_dict)

      # for mac
      #sheet.save(self.path + '/flatplates_codes.pdf')
      # for windows
      sheet.save(self.path + '\\flatplates_codes.pdf')

      # Delete files in path
      for i in qrset:
        print('deleted' + i)
        os.remove(i)

      print('Print Complete')


    # Draw two objects into label: product code and QRcode
    # The size and font changes depending on the template
    def draw_label(self, label, width, height, obj):
      if self.template_str.get() == 'Avery_6467':
        label.add(shapes.String(4, 15, obj.get('text'), fontName="Helvetica", fontSize=8))
        label.add(Image(90, 4, 30, 30, obj.get('image')))
      elif self.template_str.get() == 'Avery_5161':
        label.add(shapes.String(15, 30, obj.get('text'), fontName="Helvetica", fontSize=12))
        label.add(Image(180, 4, 60, 60, obj.get('image')))
      elif self.template_str.get() == 'Avery_94214':
        label.add(shapes.String(10, 15, obj.get('text'), fontName="Helvetica", fontSize=10))
        label.add(Image(150, 2, 40, 40, obj.get('image')))

    # return template dim -> (label_h, label_w, top_mar, bot_mar, left_mar, right_mar, num_col, num_row)
    # All measurements are in mm
    def get_template_dim(self):
      template = self.template_str.get()
      if template == 'Avery_6467':
        return (12.7, 44.45, 12.7, 10.668, 9.906, 7.874, 4, 20)
      elif template == 'Avery_94214':
        return (15.875, 76.2, 12.7, 10.668, 23.622, 7.874, 2, 16)
      elif template == 'Avery_5161':
        return (25.4, 100, 12.7, 10.668, 6.35, 7.874, 2, 10)


if __name__ == '__main__':
  root = tk.Tk()
  run = Flatplatedata(root)
  root.mainloop()