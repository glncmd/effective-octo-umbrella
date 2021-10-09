# A modification of ps-cmyk.py

# Crops all images within a folder from the bottom
# by specifying amount of pixels / percentage.

# Used for cropping out bulk photos / psd files having same watermarks at the bottom
# for direct placement in print materials e.g. yearbooks, IDs, etc.

# This script can be modified to crop from whichever side by changing Crop() method params

from comtypes.client import GetActiveObject, CreateObject
from os import listdir
from os.path import join

app = GetActiveObject("Photoshop.Application", dynamic=True)

# Folder paths
directory_folder = "C:\\Users\\Glen\\Desktop\\New folder"
destination_folder = "C:\\Users\\Glen\\Desktop\\New folder (2)"

# Save options for jpeg
jpeg_options = CreateObject("Photoshop.JPEGSaveOptions", dynamic=True)
jpeg_options.EmbedColorProfile = True
jpeg_options.FormatOptions = 1
jpeg_options.Matte = 1
jpeg_options.Quality = 12

# Save options for psd
psd_options = CreateObject("Photoshop.PhotoshopSaveOptions", dynamic=True)
psd_options.annotations = False
psd_options.alphaChannels = True
psd_options.layers = True
psd_options.spotColors = True
psd_options.embedColorProfile = True

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

for file in listdir(directory_folder):
    app.Open(join(directory_folder, file))

    current_doc = app.ActiveDocument
    
    crop(current_doc, percentage=10.5)

    current_doc.SaveAs(join(destination_folder, file), psd_options, False, None)
    current_doc.Close()
