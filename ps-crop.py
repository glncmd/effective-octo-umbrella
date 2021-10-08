# A modification of ps-cmyk.py

# Crops all images within a folder from the bottom
# by specifying amount of pixels / percentage.

# Used for cropping out bulk photos having same watermarks at the bottom
# for direct placement in print materials e.g. yearbooks, IDs, etc.

from comtypes.client import GetActiveObject, CreateObject
from os import listdir
from os.path import join

app = GetActiveObject("Photoshop.Application", dynamic=True)

# Folder paths
directoryFolder = "C:\\Users\\Glen\\Desktop\\New folder"
destinationFolder = "C:\\Users\\Glen\\Desktop\\New folder (2)"

# Save options for jpeg
jpegOptions = CreateObject("Photoshop.JPEGSaveOptions", dynamic=True)
jpegOptions.EmbedColorProfile = True
jpegOptions.FormatOptions = 1
jpegOptions.Matte = 1
jpegOptions.Quality = 12

# Save options for psd
psdOptions = CreateObject("Photoshop.PhotoshopSaveOptions", dynamic=True)
psdOptions.annotations = False
psdOptions.alphaChannels = True
psdOptions.layers = True
psdOptions.spotColors = True
psdOptions.embedColorProfile = True

def crop(doc, pixels=0, percentage=0):
    """
    Crops a Photoshop document from the bottom by
    amount of pixels or percentage of document height.
    """
    if pixels:
        doc.Crop((0, 0, doc.width, doc.height - pixels))
        return

    if percentage:
        doc.Crop((0, 0, doc.width, doc.height - (doc.height*(percentage / 100))), None, None, None)

for file in listdir(directoryFolder):
    app.Open(join(directoryFolder, file))

    currentDoc = app.ActiveDocument
    
    crop(currentDoc, percentage=10.5)

    currentDoc.SaveAs(join(destinationFolder, file), psdOptions, False, None)
    currentDoc.Close()
