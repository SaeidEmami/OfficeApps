# Office Utilities

These python scripts are intended to be called from Linux command line. 

The following is a brief description for each script:


## impressMaker.com

Creates an open office presentation (.odp) with all the images in a directory.

Images are sorted and one image is inserted in each slide.

File blankODP.zip is needed to be in the same directory as the script as it provies a barebone presentation.

How to run it from command line:
  impressMaker.py dir=imgDir [image_size=wxh][main_title=title0][sub_title1=title1][sub_title2=title2][first_page_number=n]
                             
title0, title1 and title2 can be a string or a file containing page titles in each line.

The commnad line options are not intended to replace any task that can be reproduced by modifying the master slide.





