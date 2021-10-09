# Converts a folder of RGB images / psd files to CMYK, 300dpi

# Used for preparing print-ready photos from a design project in bulk.

from comtypes.client import GetActiveObject, CreateObject
from os import listdir
from os.path import join

app = GetActiveObject("Photoshop.Application", dynamic=True)

# Folder paths
directory_folder = "C:\\Users\\Glen\\Desktop\\New folder"
destination_folder = "C:\\Users\\Glen\\Desktop\\New folder (2)"

# Color profiles
RGB = "sRGB IEC61966-2.1"
CMYK = "U.S. Web Coated (SWOP) v2"

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

def resize(doc, resolution):
    # No resampling
    doc.ResizeImage(None, None, resolution, 1)

def convert(doc, profile):
    doc.ConvertProfile(profile, 3, True, True)

for file in listdir(directory_folder):
    app.Open(join(directory_folder, file))

    current_doc = app.ActiveDocument
    
    resize(current_doc, 300)
    convert(current_doc, CMYK)

    current_doc.SaveAs(join(destination_folder, file), psd_options, False, None)
    current_doc.Close()
