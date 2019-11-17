import os
import pandas as pd
import qrcode
import csv

# Import QR Code library



def list_files(dir):
    r = []
    for root, dirs, files in os.walk(dir):
        for name in files:
            # Strip the file extension
            r.append(os.path.splitext(name)[0])
    return r

# master_list = []
pics = []

d = ["D:/Pictures_Latest/Set_11/JPEG/"]
'''

d = ["Q:/Nguyen Graphic Designer Work/Nguyen Work 1 - 139 items",
     "Q:/Nguyen Graphic Designer Work/Nguyen Set 2 - 92 items",
     "Q:/Nguyen Graphic Designer Work/Nguyen Set 3 - 219 items",
     "Q:/Nguyen Graphic Designer Work/Nguyen Set 4 - 105 items",
     "Q:/Nguyen Graphic Designer Work/Nguyen Set 5 - 83 items",
     "Q:/Nguyen Graphic Designer Work/Set 6 - 76 items",
     "Q:/Nguyen Graphic Designer Work/Set 7 - 74 items",
     "Q:/Nguyen Graphic Designer Work/Set 8 - 220 items",
     "Q:/Nguyen Graphic Designer Work/Set 9 - 726 items",
     "Q:/Nguyen Graphic Designer Work/Set 10 - same set 11",
     "Q:/Nguyen Graphic Designer Work/Set 11 - 133 photos",
     "Q:/Nguyen Graphic Designer Work/Set 12 - 43 photos",
     "Q:/Nguyen Graphic Designer Work/Set 13 - 105 photos",
     "Q:/Nguyen Graphic Designer Work/Set 14 - 234 photos",
     "Q:/Nguyen Graphic Designer Work/SET 15/JPEG/Without Watermark",
     "Q:/Nguyen Graphic Designer Work/SET 16 - 415 photos/JPEG/WithoutWatermark",
     "Q:/Nguyen Graphic Designer Work/set 18 - 48 photos/JPEG/Without Watermark",
     "C:/Users/Dinh/Desktop/Website_Pictures",
     "C:/Users/Dinh/Desktop/Online_images"]

'''
print("Depth of Directory", len(d))
# Traverse each directory to get filenames
for i in range(0, len(d)):
    pics = pics + list_files(d[i])

# Remove duplicates
pics = list(set(pics))

# Sort
pics.sort()

df = pd.DataFrame(pics)
df.to_csv('./QR/Out/new_qr.csv', index=False)

print(pics)


#Start of QRcode creation

with open('./QR/Out/new_qr.csv', 'r') as f:
    reader = csv.reader(f)
    item_list = list(reader)


for item in item_list:
    # Create qr code instance
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=7,
        border=2,
    )
    # Add data
    qr.add_data(item[0])
    qr.make(fit=True)
    # Create an image from the QR Code instance
    img = qr.make_image()
    # Save it somewhere, change the extension as needed:
    # img.save("image.png")
    # img.save("image.bmp")
    # img.save("image.jpeg")
    img.save("./QR/New/qr_" + item[0]+".png")