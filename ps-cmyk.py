# Converts a folder of RGB images to CMYK, 300dpi
# Used for preparing print-ready photos from a design project in bulk.

from comtypes.client import GetActiveObject, CreateObject
from os import listdir
from os.path import join

app = GetActiveObject("Photoshop.Application", dynamic=True)

# Folder paths
directoryFolder = "C:\\Users\\Glen\\Desktop\\New folder"
destinationFolder = "C:\\Users\\Glen\\Desktop\\New folder (2)"

# Color profiles
RGB = "sRGB IEC61966-2.1"
CMYK = "U.S. Web Coated (SWOP) v2"

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

def resize(doc, resolution):
    # No resampling
    doc.ResizeImage(None, None, resolution, 1)

def convert(doc, profile):
    doc.ConvertProfile(profile, 3, True, True)

for file in listdir(directoryFolder):
    app.Open(join(directoryFolder, file))

    currentDoc = app.ActiveDocument
    
    resize(currentDoc, 300)
    convert(currentDoc, CMYK)
    
    currentDoc.SaveAs(join(destinationFolder, file), psdOptions, False, None)
    currentDoc.Close()
