import pyqrcode
import qrcode
from PIL import Image, ImageDraw, ImageFont

start = 1
num_sheets = 1
label_text = "This Is Sample Label Text"

count = start - 1

for j in range(1, num_sheets + 1):
    break
    sheet = Image.new('RGBA', (563, 750), (255, 255, 255, 255))

    for i in range(0, 80):
        count = count + 1
        seq = str(count).zfill(7)
        qr = QRCode('https://www.physics.purdue.edu/cmsfpix/CMSTrackerFlatSheetsDB_Data/C4-76-12345-6-P21-1/C3-76-12345-6-P21-1.pdf')
        qr.png('qr.png', scale=1)

        # Open qr image and dimensions
        qr_img = Image.open('qr.png', 'r')
        qr_img = qr_img.resize((35, 35))
        qr_img_w, qr_img_h = qr_img.size

        # Bigger Canvas image that label goes onto
        #label = Image.new('RGBA', (288, 72), (255, 255, 255, 255))
        label = Image.new('RGBA', (126, 36), (255, 255, 255, 255))
        bg_w, bg_h = label.size
        offset = ((bg_h - qr_img_h) // 2, (bg_h - qr_img_h) // 2)
        label.paste(qr_img, offset)

        # write text
        d = ImageDraw.Draw(label)
        label_h_offset = int(155 - (5 * len(label_text) / 2))
        #d.text((label_h_offset, 22), label_text, fill=(0, 0, 0))
        d.text((10, 10), label_text, fill=(0, 0, 0))

        #h_offset = int(i % 2 * 302)
        h_offset = int(i % 2 * 150)
        row = int(abs(i / 4))
        #y_offset = int(row * 72)
        y_offset = int(row * 36)
        sheet.paste(label, (h_offset, y_offset))

    sheet_name = 'sheet' + str(j) + '.png'
    sheet.save(sheet_name)

url = 'https://www.physics.purdue.edu/cmsfpix/CMSTrackerFlatSheetsDB_Data/C3-76-12345-6-P21-1/C3-76-12345-6-P21-1.pdf'
qrcode = pyqrcode.create(url)
tmp = 'generated_qr/' + str('C3-76-12345-6-P21-1') + '.png'
qrcode.png(tmp, scale=1)