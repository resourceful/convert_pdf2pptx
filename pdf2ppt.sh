#!/usr/bin/env python3.10
import sys, os

from pdf2image import convert_from_path
from io import BytesIO
from pptx import Presentation

if (len(sys.argv) != 2 ):
	print("usage: ", sys.argv[0], "PDF-FILE-NAME")
	exit()

pdf = sys.argv[1]
base = pdf.split(".pdf")[0]

# create blank slide
ppt = Presentation()
core = ppt.core_properties
core.author = "Seongjin Lee"
core.title = base
core.subject = "Technical reports"
core.last_modified_by = "Seongjin Lee"
core.comments = "PPTX generated from PDF"

# Convert PDF to list of images
print("\nConverting", pdf, "in progress\n(time depends on no. of pages)")
pages = convert_from_path(pdf, 150, fmt='ppm', thread_count=8)
print("Conversion complete.\n")

#Append slides
for i, img in enumerate(pages):
	print("Processing slide", str(i+1), "of", len(pages), end="\r")
	imgfile = BytesIO()
	img.save(imgfile, format='png')
	width, height = img.size

	# Slide dimension (pptx default dimension): 
	# wide (12192000x6858000), 4x3 (9144000x6858000)
	ppt.slide_height = 6858000
	ppt.slide_width = 12192000
	
	# Add slide (add blank slide template)
	slide = ppt.slides.add_slide(ppt.slide_layouts[6]) 
	pic = slide.shapes.add_picture(imgfile, 0, 0, height=ppt.slide_height)

# Save Powerpoint presentation
ppt.save(base + '.pptx')
print("\n\n", base + '.pptx', " created.", sep="")
