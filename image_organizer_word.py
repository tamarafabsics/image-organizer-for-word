#Image Organizer for Word


print ("Hi!")
folder = input ("Which images would you like to insert into a Word document? Enter the path to the folder! ").strip( )

docname = input ("What should the document be named? ")



#Create a list with the images

import os
files = os.listdir(folder)


images = []

imageformat = (".jpg", ".png", ".jpeg")

for image in files:
    if image.lower().endswith(imageformat):
        image = folder + "/" + image
        images.append(image)
    else:
        print ("Images with invalid image format! :(")
        import sys
        sys.exit()


def sortkey(image):
    return os.path.getctime(image)

images.sort(key=sortkey)



#Create a Word document

import docx
from docx.shared import Cm

 
doc = docx.Document()


from PIL import Image
from PIL import ImageOps

def image_turn(imagepath):
    imagetoturn = Image.open(imagepath)
    turnedimage = ImageOps.exif_transpose(imagetoturn)  
    turnedimagepath = imagepath + "_turned.jpg"
    turnedimage.save(turnedimagepath)
    return turnedimagepath

imageheight = float(input("What height should the images have? (in cm) "))

for imagecorrect in images:
    imageturned = image_turn(imagecorrect)
    doc.add_picture(imageturned, height=Cm(imageheight))
    os.remove(imageturned)


doc.save(os.path.join(folder, f"{docname}.docx"))
print(f"The document '{docname}.docx' has been saved to: {folder}")
