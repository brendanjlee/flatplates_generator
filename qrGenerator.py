from __future__ import print_function
import tkinter as tk # GUI
import subprocess
from __main__ import *
import lxml
from urllib import parse
from mailmerge import MailMerge # for writing to WORD.doc
import pyqrcode
import png

template = "templates/1_5x13_4template.docx"
document = MailMerge(template)

template_options = ['Avery 6467', 'Avery 94214', 'Avery 5161']
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

    # Function initializes the user interface
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
        self.quantlab = tk.Label(self.parent,  text="Range of plates:", fg='black', relief="sunken", font=("Helvetica", "12", "bold"))
        self.quantlab.place(x=70, y=260)
        self.quant_low=tk.Entry(self.parent)
        self.quant_low.place(x=300, y=260, width=50)

        self.quant_high = tk.Entry(self.parent)
        self.quant_high.place(x=355, y=260, width=50)

        # Generate Button
        global prevButton
        prevButton = tk.Button(self.parent, text="Generate Labels", fg='black', bg='green', relief="ridge", font=("Helvetica", "14", "bold"), command=self.Preview)
        prevButton.place(x=175, y=435)

    # Generates preview of the code
    def Preview(self):
        fib = self.fibtype_str.get()
        gsm = self.gsmtype_str.get()
        lot = self.lottype.get()
        manyear = self.manyearlab_str.get()
        quant_lo = self.quant_low.get()
        quant_hi = self.quant_high.get()

        x = int(quant_lo)
        quantset = range(1, x + 1, 1)
        codeset = [None] * x
        qrset = [None] * x
        i = 0
        #LOOPS CODES INTO SET OF LEN = QAUNTITY
        while x > 0 and i < (int(quant)):
            for n in quantset:
                # Generate part code
                code = "{}-{}-{}-{}-{}".format(fib, gsm, lot, manyear, n)
                codeset[i] = code
                # Generate QRcode and store it in /generated_qr/
                url = generate_url(code)
                qrcode = pyqrcode.create(url)
                tmp_png = "generated_qr/" + str(code) + ".png"
                qrcode.png(tmp_png, scale=4)
                qrset[i] = tmp_png
                i += 1

        # Print Preview on GUI
        preview = tk.Label(self.parent, text="Preview:", fg='black', font=("Helvetica", "10", "bold"))  # preview label
        preview.place(x=75, y=340)
        codePrev = tk.Label(self.parent, text=code, fg='black', relief="sunken", font=("Helvetica", "24", "bold")) #product code label
        codePrev.place(x=70, y=370)

        print(document.get_merge_fields())
        # Write to Document
        #QRurl = 'C:\\Users\\jackg\\PycharmProjects\\flatplates\\QRs\\Flatplates.jpg'
        #reformURL = urllib.parse.quote_plus(QRurl)
        dictOfCodes = {str(i): codeset[i] for i in range(0,int(quant))}
        #dictOfCodes['imagepath'] = 'generated_qr/'
        document.merge(**dictOfCodes)
        #document.merge(url=code)
        document.write('generated_docx/test.docx')
        #print(dictOfCodes)

if __name__ == '__main__':
    root = tk.Tk()
    run = Flatplatedata(root)
    root.mainloop()