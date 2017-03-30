#!C:\Python27
# -*- coding: utf-8
import json
import os
from PIL import Image
import re
import sys
import shutil
import xlwt
import time
from random import random
from clint.textui import progress
Image.MAX_IMAGE_PIXELS = None
sys.path.insert(0, os.path.abspath('..'))


__author__ = "Alberto Punter"
__copyright__ = "Copyright 2017, TheCropper"
__license__ = "GNU GENERAL PUBLIC LICENSE"
__version__ = "1.0"
__email__ = "dragost11@gmail.com"



# =============================================================================
# Configuration
# =============================================================================
default_quality_val = 90
input_folder = './Input/'
date = time.strftime("%d-%m-%Y_%H-%M")
output_folder = './Output/'
with open('resolutions.json') as data_file:
    resAndPath = json.load(data_file)
fileType = '.jpg'
random_str = "zwEwCgHkRZ"
duplicate = False
organize = False
head_style = xlwt.easyxf('font: name Arial, colour black, bold on')
dpi = 72


# =============================================================================
# Arrays
# =============================================================================
lstFiles = []
dupFiles = []
skus = []
numPhotos = {}
lstDir = os.walk(input_folder)



# =============================================================================
# Functions
# =============================================================================
def ensure_dir(f):
    d = os.path.dirname(f)
    if not os.path.exists(d):
        os.makedirs(d)


def fcount(path):
    count = 0
    for f in os.listdir(input_folder):
        if os.path.isfile(os.path.join(input_folder, f)):
            count += 1
    return count



# =============================================================================
# First Open
# =============================================================================
if not os.path.exists(input_folder) and not os.path.exists(output_folder):
    print "Iniciando por primera vez."
    ensure_dir(input_folder)
    ensure_dir(output_folder)
    print "Introduzca las photoes a recortar en " + input_folder
    print "Saliendo..."
    time.sleep(10)
    sys.exit()




# Set sku_lenght
sku_lenght = raw_input("Number of digits of SKU?: ")
try:
    sku_lenght = int(sku_lenght)
    pattern = re.compile('([0-9]{' + str(sku_lenght) + '})')
except ValueError:
    print("That's not an int!")
    raw_input("")
    sys.exit()



# DUplicate Image for variuos SKUs
if raw_input("The images of the products have "
             "several SKUs in the name? Y/N: ").upper() == "Y":
    duplicate = True


# Organize by SKU number
if raw_input("Organize photos by sku? Y/N: ").upper() == "Y":
    organize = True


print "Starting process..."


path, dirs, files = lstDir.next()
# Create folders
if len(files) == 0:
    print "Introduzca las photos para recortar en " + input_folder
    print "Saliendo..."
    time.sleep(10)
    sys.exit()
else:
    output_folder += date + "/"
    ensure_dir(output_folder)
    for row in resAndPath:
        ensure_dir(output_folder + row['folder'] + "/")


    # Rename to avoid duplicates
    with progress.Bar(label="Revision: ",
                      expected_size=fcount(input_folder)
                      ) as bar:
        val = 1
        lstDir = os.walk(input_folder)
        for root, dirs, files in lstDir:
            for fichero in files:
                (fileName, extension) = os.path.splitext(fichero)
                os.rename(input_folder + fileName + extension,
                          input_folder + fileName + random_str + extension)
                bar.show(val)
                val += 1


    # Rename files
    with progress.Bar(label="Rename: ",
                      expected_size=fcount(input_folder)
                      ) as bar:
        val = 1
        lstDir = os.walk(input_folder)
        for root, dirs, files in lstDir:
            for fichero in files:
                (fileName, extension) = os.path.splitext(fichero)
                if(extension.lower() == ".jpg" or
                   extension.lower() == ".png" or
                   extension.lower() == ".tif"):

                    if duplicate:
                        matcher = pattern.findall(fileName)
                    else:
                        busqueda = pattern.search(fileName).group(1)
                        matcher = []
                        matcher.append(busqueda)
                    if len(matcher) > 0:
                        for m in matcher:
                            nuevoNombre = m
                            i = 1
                            while (nuevoNombre + extension in lstFiles):
                                nuevoNombre = "{}_{}".format(
                                        nuevoNombre[:sku_lenght],
                                        str(i)
                                    )
                                i += 1

                            try:
                                shutil.copy(input_folder +
                                            fileName +
                                            extension,
                                            input_folder +
                                            nuevoNombre +
                                            extension)
                            except shutil.Error as e:
                                print('Error: %s' % e)
                            else:
                                lstFiles.append(nuevoNombre + extension)

                        os.remove(input_folder + fileName + extension)
                bar.show(val)
                val += 1



    # Resize images
    with progress.Bar(label="Resize: ", expected_size=len(lstFiles)) as bar:
        val = 1
        for nimage in lstFiles:

            # Add SKU to Unique SKUS list
            if nimage[:sku_lenght] not in skus:
                skus.append(nimage[:sku_lenght])

            # Count number images
            if nimage[:sku_lenght] not in numPhotos.keys():
                numPhotos[nimage[:sku_lenght]] = 1
            else:
                numPhotos[nimage[:sku_lenght]] += 1

            for res in resAndPath:
                photo = Image.open(input_folder + nimage)
                size = res['width'], res['height']

                # More Weight
                if photo.size[0] > photo.size[1]:
                    new_height = size[0] * photo.size[1] / photo.size[0]
                    photo = photo.resize(
                        (size[0], new_height),
                        Image.ANTIALIAS
                    )
                # More Height
                else:
                    new_width = size[1] * photo.size[0] / photo.size[1]
                    photo = photo.resize(
                        (new_width, size[1]),
                        Image.ANTIALIAS
                    )

                background = Image.new('RGB', size, (255, 255, 255, 0))
                background.paste(photo,
                                 ((size[0] - photo.size[0]) / 2,
                                  (size[1] - photo.size[1]) / 2))
                icc_profile = background.info.get("icc_profile")

                background.save(
                    "{}{}/{}{}{}{}".format(
                        output_folder,
                        res['folder'],
                        nimage[:sku_lenght],
                        res['sufix'],
                        nimage[sku_lenght:][:-4],
                        fileType
                    ),
                    'JPEG',
                    quality=default_quality_val,
                    icc_profile=icc_profile,
                    dpi=(dpi, dpi)
                )

            bar.show(val)
            val += 1


sys.exit()




if organize:
    print "Mmm Working on it... ^^"




# Clean
for res in resAndPath:
    shutil.rmtree(res[3], ignore_errors=False, onerror=None)




wb = xlwt.Workbook()
ws = wb.add_sheet('Page1', cell_overwrite_ok=True)
ws.write(0, 0, 'SKU', head_style)
ws.write(0, 1, 'Num Photos', head_style)

# Write (file, column, data)
for num, sku in enumerate(skus):
    ws.write(num + 1, 0, sku)
    ws.write(num + 1, 1, numPhotos[sku])


wb.save('data_generated.xls')
