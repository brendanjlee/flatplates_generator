
from __future__ import print_function
import tkinter as tk # GUI
from tkinter.filedialog import asksaveasfile
import subprocess
from __main__ import *
#import lxml
from urllib import parse
from mailmerge import MailMerge # for writing to WORD.doc
import pyqrcode
import png
import labels
from reportlab.graphics import shapes
from reportlab.graphics.shapes import Image
import qrcode
from PIL import Image as pil_image

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
        self.template_str = tk.StringVar(self.parent)
        self.template_str.set(template_options[0])
        self.template_type = tk.OptionMenu(self.parent, self.template_str, *template_options)
        self.template_type.place(x=300, y=38)

        # Label and Field for Fiber Type
        self.fiblab = tk.Label(self.parent, text="Fiber type & ply #:", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.fiblab.place(x=70, y=80)
        self.fibtype_str = tk.StringVar(self.parent)
        self.fibtype_str.set(fiber_options[0])
        self.fibtype = tk.OptionMenu(self.parent, self.fibtype_str, *fiber_options)
        self.fibtype.place(x=300, y=83)

        # Label and Field GSM
        self.gsmlab=tk.Label(self.parent,text="GSM:",fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.gsmlab.place(x=70, y=125)
        self.gsmtype_str = tk.StringVar(self.parent)
        self.gsmtype_str.set(gsm_options[0])
        self.gsmtype = tk.OptionMenu(self.parent, self.gsmtype_str, *gsm_options)
        self.gsmtype.place(x=300, y=128)

        # Label and Field Lot
        self.lotlab = tk.Label(self.parent, text="Lot #:", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.lotlab.place(x=70, y=170)
        self.lottype=tk.Entry(self.parent)
        self.lottype.insert(0, "#####-#")
        self.lottype.place(x=300, y=173)

        # Label and Field for Manufacture Year
        self.manyearlab = tk.Label(self.parent, text="Manufacturer & Year:", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.manyearlab.place(x=70, y=215)
        self.manyearlab_str = tk.StringVar(self.parent)
        self.manyearlab_str.set(manyear_options[0])
        self.manyeartype = tk.OptionMenu(self.parent, self.manyearlab_str, *manyear_options)
        self.manyeartype.place(x=300, y=218)

        # Label and Field for Part Identifier
        self.quantlab = tk.Label(self.parent,  text="Range of plates (start, end):", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.quantlab.place(x=70, y=260)
        # Minimum quant field
        self.quant_low = tk.Entry(self.parent)
        self.quant_low.insert(0, "1")
        self.quant_low.place(x=300, y=260, width=50)
        # Maximum quant field
        self.quant_high = tk.Entry(self.parent)
        self.quant_high.insert(0, "1")
        self.quant_high.place(x=355, y=260, width=50)

        # Generate Button
        global prevButton
        prevButton = tk.Button(self.parent, text="Generate Labels", fg='black', bg='green', relief="ridge", font=("Helvetica", "14", "bold"), command=self.generate_codes)
        prevButton.place(x=175, y=435)

    # Generates preview of the code
    # Also creates a list of codes based on quantity
    def generate_codes(self):
        #data = [('All tyes(*.*)', '*.*')]
        #file = asksaveasfile(filetypes = data, defaultextension = data)

        fib = self.fibtype_str.get()
        gsm = self.gsmtype_str.get()
        lot = self.lottype.get()
        manyear = self.manyearlab_str.get()
        quant_lo = 1
        quant_lo = self.quant_low.get()
        quant_hi = 1
        quant_hi = self.quant_high.get()
        template = self.template_str.get()

        print(template)

        x = int(quant_hi)
        quantset = range(1, x + 1, 1)
        print(type(quantset))
        codeset = []
        qrset = []
        code = ''   # stores sheet string

        for i in range(int(quant_lo), int(quant_hi) + 1):
            # Genreate sheet string code
            code = "{}-{}-{}-{}-{}".format(fib, gsm, lot, manyear, i)
            codeset.append(code)
            url = generate_url(code)
            # Generate QR code
            qrcode = pyqrcode.create(url)
            tmp_png = "generated_qr/" + str(code) + ".png"
            qrcode.png(tmp_png, scale=2)
            qrset.append(tmp_png)

        # Print Preview on GUI
        preview = tk.Label(self.parent, text="Preview:", fg='black', font=("Helvetica", "10", "bold"))  # preview label
        preview.place(x=75, y=340)
        codePrev = tk.Label(self.parent, text=code, fg='black', relief="sunken", font=("Helvetica", "24", "bold")) #product code label
        codePrev.place(x=70, y=370)

        # Name for Word Document to be printed
        #self.outfile_name = "{}-{}-{}-{}_{}.docx".format(fib, gsm, lot, manyear, self.template_str.get())
        self.get_pdf(codeset, qrset, template)
        #self.print_labels(codeset, qrset, template)

    # All sizes are in mm
    def get_pdf(self, codeset, qrset, template):
        #
        label_h = 12.7
        label_w = 44.45
        top_mar = 12.7
        bot_mar = 10.668
        left_mar = 9.906
        right_mar = 7.874
        num_col = 4
        num_row = 20

        (label_h, label_w, top_mar, bot_mar, left_mar, right_mar, num_col, num_row) = self.get_template(template)

        # specs for avery 6467
        specs = labels.Specification(215.9, 279.4, num_col, num_row, label_w, label_h, corner_radius=1, top_margin=top_mar, bottom_margin=bot_mar, left_margin=left_mar, right_margin=right_mar)
        sheet = labels.Sheet(specs, self.draw_label, border=True)

        #qr = pyqrcode.create('https://www.physics.purdue.edu/cmsfpix/CMSTrackerFlatSheetsDB_Data/C4-76-12345-6-P21-1/C3-76-12345-6-P21-1.pdf')
        #qr.png('qr.jpg', scale=1)

        qr = qrcode.QRCode()


        # print the labels
        j = 0
        for i in range(len(codeset)):
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4)
            qr.add_data('There is a lot of data')

            _print_dict = {
                'text' : codeset[i],
                'image' : 'testqr.png',
            }

            sheet.add_label(_print_dict)
            j += 1



        sheet.save('output.pdf')
        print('Printing Complete')

    def get_template(self, template):
        if template == 'Avery_6467':
            # return (label_h, label_w, top_mar, bot_mar, left_mar, right_mar, num_col, num_row)
            return (12.7, 44.45, 12.7, 10.668, 9.906, 7.874, 4, 20)
        elif template == 'Avery_94214':
            # return (label_h, label_w, top_mar, bot_mar, left_mar, right_mar, num_col, num_row)
            return (15.875, 76.2, 12.7, 10.668, 21.336, 7.874, 2, 16)
        elif template == 'Avery_5161':
            return (25.4, 101.6, 12.7, 10.668, 6.35, 7.874, 2, 10)

    def draw_label(self, label, width, height, obj):
        label.add(shapes.String(2, 6, obj.get('text'), fontName="Helvetica", fontSize=8))
        label.add(Image(70 + len(obj.get('text')), 2, 30, 30, obj.get('image')))




if __name__ == '__main__':
    root = tk.Tk()
    run = Flatplatedata(root)
    root.mainloop()