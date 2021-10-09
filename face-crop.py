# Given multiple pictures of persons, this script searches
# for the location of a face in each one, and crops it to the
# framing and dimensions of an ID picture.

# This saves time and effort by automating the task of manually cropping a
# collection of pictures with different framing, sizes, formats, and resolutions.
# This is effective, for example, in ID card projects.

import face_recognition as fr
import os
from PIL import Image

# Folder paths
directory_folder = "C:\\Users\\Glen\\Desktop\\New folder"
destination_folder = "C:\\Users\\Glen\\Desktop\\New folder (2)"

def get_images(folder):
    '''
    lists pngs and jpgs in a folder for processing
    '''
    image_paths = []
    for dirpath, dnames, fnames in os.walk(folder):
        for file in fnames:
            if file.endswith('.jpg') or file.endswith('.png'):
                image_paths.append(file)
    
    return image_paths

process_images = get_images(directory_folder)

image_count = 0

for image in process_images:
    # Locate face coordinates
    current_image = fr.load_image_file(os.path.join(directory_folder, image))
    faceLoc = fr.face_locations(current_image)
    (top, right, bottom, left) = faceLoc[0]

    # Compute padding
    pad = (right - left) / 2
    left = left - pad
    right = right + pad
    top = top - pad
    bottom = bottom + pad

    # Adjust padding if it exceeds image bounds
    pil_image = Image.open(os.path.join(directory_folder, image))
    if left < 0: left = 0
    if right > pil_image.width: right = pil_image.width
    if top < 0: top = 0
    if bottom > pil_image.height: bottom = pil_image.height

    # Crop image
    crop_image = pil_image.crop((left, top, right, bottom))

    # Save to destination folder
    crop_image.save(os.path.join(destination_folder, image))

    image_count += 1

print(f"{image_count} images processed")
