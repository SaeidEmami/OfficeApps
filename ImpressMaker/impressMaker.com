#!/usr/bin/python

###################################################################################################
# Available for use under the [MIT License](http://en.wikipedia.org/wiki/MIT_License)             #
# Writer: Saeid Emami                                                                             #
###################################################################################################

'''
Creates an open office presentation (.odp) with all the images in a directory.

Images are sorted and one image is inserted in each slide.

File blankODP.zip is needed to be in the same directory as the script as it provies a barebone presentation.

How to run it from command line:
  impressMaker.com dir=imgDir [image_size=wxh][image_position=posXxposY][main_title=title0][sub_title1=title1][sub_title2=title2][first_page_number=n]
                             
image_size and image_position are in cm.
title0, title1 and title2 can be a string or a file containing page titles in each line.

The commnad line options are not intended to replace any task that can be reproduced by modifying the master slide.
'''


def makeImpress(img_dir, size, position, title_list0, title_list1, title_list2, page_n):

    def xml_per_image(image, posx, posy, w, h, counter, title0, title1, title2, page_n):
        '''
        Output open-office xml content for each image.
        ''' 
        result = '''<draw:page draw:name="pageCOUNTER" draw:style-name="dp1" draw:master-page-name="Default">'''
        result += '''<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" svg:width="%icm" svg:height="%icm" svg:x="%icm" svg:y="%icm"><draw:image xlink:href="Pictures/%s" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><text:p/></draw:image></draw:frame>''' %(w, h, posx, posy, image)
        if title0 is not None:
            result += '''<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" svg:width="21.5cm" svg:height="2cm" svg:x="3.3cm" svg:y="1.1cm"><draw:text-box><text:p text:style-name="P1"><text:span text:style-name="T1">%s</text:span></text:p></draw:text-box></draw:frame>''' %title0
        if title1 is not None:
            result += '''<draw:frame draw:style-name="gr5" draw:text-style-name="P4" draw:layer="layout" svg:width="23.5cm" svg:height="1.1cm" svg:x="2.5cm" svg:y="3.4cm"><draw:text-box><text:p text:style-name="P3"><text:span text:style-name="T2">%s</text:span></text:p></draw:text-box></draw:frame>''' %title1
        if title2 is not None:
            result += '''<draw:frame draw:style-name="gr4" draw:text-style-name="P1" draw:layer="layout" svg:width="21cm" svg:height="1.5cm" svg:x="3.6cm" svg:y="18cm"><draw:text-box><text:p text:style-name="P1">%s</text:p></draw:text-box></draw:frame>''' %title2
        if page_n is not None:
            result += '''<draw:frame draw:style-name="gr5" draw:text-style-name="P5" draw:layer="layout" svg:width="3cm" svg:height="1.1cm" svg:x="12.5cm" svg:y="19.7cm"><draw:text-box><text:p text:style-name="P5"><text:span text:style-name="T3">%i</text:span></text:p></draw:text-box></draw:frame>''' %page_n
        result += '''<presentation:notes draw:style-name="dp2"><draw:page-thumbnail draw:style-name="gr3" draw:layer="layout" svg:width="13.968cm" svg:height="10.476cm" svg:x="3.6cm" svg:y="2.123cm" draw:page-number="1" presentation:class="page"/><draw:frame presentation:style-name="pr1" draw:layer="layout" svg:width="17.271cm" svg:height="12.572cm" svg:x="2.159cm" svg:y="13.271cm" presentation:class="notes" presentation:placeholder="true"><draw:text-box/></draw:frame></presentation:notes></draw:page>'''
        return result


    import os
    import shutil
    import zipfile
    import errno

    if img_dir is None:
        img_dir = "."

    try:
        img_width, img_height = map(int, size.split('x'))
    except:
        img_width, img_height = 17.0, 12.0

    try:
        img_x, img_y = map(int, location.split('x'))
    except:
        img_x, img_y = 6.0, 5.0

    try:
        page_init = int(page_n)
    except:
        page_init = None

    barebone_f = "blankODP.zip"
    if not os.path.isfile(barebone_f):
        print("Barebone repository file is missing. Please move " + barebone_f + " to the current directory.")
        return False
    barebone_zf = zipfile.ZipFile(barebone_f, 'r')
    
    rep_d = os.path.join(img_dir, "_tempDir")
    try:
        os.mkdir(rep_d)
    except OSError as e:
        if e.errno == errno.EEXIST:
            print("Temporary repository directory already exists.")
        else:
            print("Failed to create repository directory " + rep_d + ".")
        return

    try:
        os.mkdir(os.path.join(rep_d, "Configurations2"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "accelerator"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "floater"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "images"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "menubar"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "popupmenu"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "progressbar"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "statusbar"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "toolbar"))
        os.mkdir(os.path.join(rep_d, "Configurations2", "toolpanel"))
        os.mkdir(os.path.join(rep_d, "META-INF"))
        os.mkdir(os.path.join(rep_d, "Pictures"))
        os.mkdir(os.path.join(rep_d, "Thumbnails"))
    except:
        print("Failed to create a subdirectory in " + rep_d + ".")
        return

    content_initial = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:officeooo="http://openoffice.org/2009/office" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2"><office:scripts/><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:background-visible="true" presentation:background-objects-visible="true" presentation:display-footer="true" presentation:display-page-number="false" presentation:display-date-time="true"/></style:style><style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties presentation:display-header="true" presentation:display-footer="true" presentation:display-page-number="false" presentation:display-date-time="true"/></style:style><style:style style:name="gr1" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-horizontal-align="center" draw:textarea-vertical-align="middle" draw:color-mode="standard" draw:luminance="0%" draw:contrast="0%" draw:gamma="100%" draw:red="0%" draw:green="0%" draw:blue="0%" fo:clip="rect(0cm, 0cm, 0cm, 0cm)" draw:image-opacity="100%" style:mirror="none"/></style:style><style:style style:name="gr2" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:stroke="none" svg:stroke-color="#000000" draw:fill="none" draw:fill-color="#ffffff" draw:auto-grow-height="true" draw:auto-grow-width="false" fo:max-height="0cm" fo:min-height="1.75cm"/></style:style><style:style style:name="gr3" style:family="graphic"><style:graphic-properties style:protect="size"/></style:style><style:style style:name="gr4" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:stroke="none" svg:stroke-color="#000000" draw:fill="none" draw:fill-color="#ffffff" draw:auto-grow-height="true" draw:auto-grow-width="false" fo:max-height="0cm" fo:min-height="1.25cm"/></style:style><style:style style:name="gr5" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:stroke="none" svg:stroke-color="#000000" draw:fill="none" draw:fill-color="#ffffff" draw:auto-grow-height="true" draw:auto-grow-width="false" fo:max-height="0cm" fo:min-height="0.85cm"/></style:style><style:style style:name="pr1" style:family="presentation" style:parent-style-name="Default-notes"><style:graphic-properties draw:fill-color="#ffffff" draw:auto-grow-height="true" fo:min-height="12.572cm"/></style:style><style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style><style:style style:name="P2" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/><style:text-properties fo:font-size="22pt" fo:font-weight="bold" style:font-size-asian="22pt" style:font-weight-asian="bold" style:font-size-complex="22pt" style:font-weight-complex="bold"/></style:style><style:style style:name="P3" style:family="paragraph"><style:paragraph-properties fo:text-align="start"/><style:text-properties fo:font-size="20pt" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" style:font-size-asian="20pt" style:font-size-complex="20pt"/></style:style><style:style style:name="P4" style:family="paragraph"><style:paragraph-properties fo:text-align="start"/><style:text-properties fo:font-size="20pt" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" fo:font-weight="bold" style:font-size-asian="20pt" style:font-weight-asian="bold" style:font-size-complex="20pt" style:font-weight-complex="bold"/></style:style><style:style style:name="P5" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/><style:text-properties fo:font-size="14pt" style:text-underline-style="none" fo:font-weight="normal" style:font-size-asian="14pt" style:font-weight-asian="normal" style:font-size-complex="14pt" style:font-weight-complex="normal"/></style:style><style:style style:name="T1" style:family="text"><style:text-properties fo:font-size="24pt" fo:font-weight="bold" style:font-size-asian="22pt" style:font-weight-asian="bold" style:font-size-complex="22pt" style:font-weight-complex="bold"/></style:style><style:style style:name="T2" style:family="text"><style:text-properties fo:font-size="20pt" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" fo:font-weight="bold" style:font-size-asian="20pt" style:font-weight-asian="bold" style:font-size-complex="20pt" style:font-weight-complex="bold"/></style:style><style:style style:name="T3" style:family="text"><style:text-properties fo:font-size="14pt" style:text-underline-style="none" fo:font-weight="normal" style:font-size-asian="14pt" style:font-weight-asian="normal" style:font-size-complex="14pt" style:font-weight-complex="normal"/></style:style><text:list-style style:name="L1"><text:list-level-style-bullet text:level="1"><style:list-level-properties text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="2"><style:list-level-properties text:space-before="0.6cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="3"><style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="4"><style:list-level-properties text:space-before="1.8cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="5"><style:list-level-properties text:space-before="2.4cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="6"><style:list-level-properties text:space-before="3cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="7"><style:list-level-properties text:space-before="3.6cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="8"><style:list-level-properties text:space-before="4.2cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="9"><style:list-level-properties text:space-before="4.8cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="10"><style:list-level-properties text:space-before="5.4cm" text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet></text:list-style></office:automatic-styles><office:body><office:presentation>'''

    content_final = '''<presentation:settings presentation:mouse-visible="false"/></office:presentation></office:body></office:document-content>'''

    if img_dir is None:
        img_dir = "."
    try:
        files = os.listdir(img_dir)
    except:
        print("Invalid directory for images.")
        return
    files.sort()
    file_exts = [os.path.splitext(f) for f in files]
    img_files = [f for f in file_exts if f[1].lower() in [".png", ".jpg", ".jpeg"]]

    page_number = page_init
    content = content_initial
    for i, f in enumerate(img_files):
        title0 = title_list0[i % len(title_list0)]
        title1 = title_list1[i % len(title_list1)]
        title2 = title_list2[i % len(title_list2)]
        content += xml_per_image(f[0] + f[1], img_x, img_y, img_width, img_height, i, title0, title1, title2, page_number)
        if page_init is not None:
            page_number += 1

    content += content_final

    try:
        contentFile = open(os.path.join(rep_d , "content.xml"), 'w')
    except:
        print("Failed to write in " + rep_d + ".")
        return
    contentFile.write(content)
    contentFile.close()

    try:
        f = open(os.path.join(rep_d, "meta.xml"), 'w')
    except:
        print("Failed to write in " + rep_d + ".")
        return
    f.write(barebone_zf.read("meta.xml"))
    f.close()

    try:
        f = open(os.path.join(rep_d, "mimetype"), 'w')
    except:
        print("Failed to write in " + rep_d + ".")
        return
    f.write(barebone_zf.read("mimetype"))
    f.close()

    try:
        f = open(os.path.join(rep_d, "settings.xml"), 'w')
    except:
        print("Failed to write in " + rep_d + ".")
        return
    f.write(barebone_zf.read("settings.xml"))
    f.close()

    try:
        f = open(os.path.join(rep_d, "Configurations2", "accelerator", "current.xml"), 'w')
    except:
        print("Failed to write in " + rep_d + ".")
        return
    f.write(barebone_zf.read("meta.xml"))
    f.close()


    manifest1 = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.presentation" manifest:version="1.2" manifest:full-path="/"/>
 <manifest:file-entry manifest:media-type="" manifest:full-path="Configurations2/accelerator/current.xml"/>
 <manifest:file-entry manifest:media-type="application/vnd.sun.xml.ui.configuration" manifest:full-path="Configurations2/"/>\n'''
    manifest2 = ''' <manifest:file-entry manifest:media-type="image/XXXEXTXXX" manifest:full-path="Pictures/XXXIMAGEXXX"/>\n'''
    manifest3 = ''' <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
 <manifest:file-entry manifest:media-type="" manifest:full-path="Thumbnails/thumbnail.png"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>
</manifest:manifest>'''

    manifest = manifest1
    for f in img_files:
        manifest += manifest2.replace('XXXIMAGEXXX', (f[0] + f[1])).replace('XXXEXTXXX', f[1])
    manifest += manifest3

    try:
        manifest_f = open(os.path.join(rep_d, "META-INF", "manifest.xml"), 'w')
    except:
        print("Failed to write in " + rep_d + ".")
        return
    manifest_f.write(manifest)
    manifest_f.close()

    try:
        for f in img_files:
            shutil.copy(os.path.join(img_dir, (f[0] + f[1])), os.path.join(rep_d, "Pictures") )
    except:
        print("Failed to copy images to " + rep_d + ".")
        return

    try:
        f = open(os.path.join(rep_d, "Thumbnails", "thumbnail.png"), 'w')
    except:
        print("Failed to write in " + rep_d + ".")
        return
    f.write(barebone_zf.read("thumbnail.png"))
    f.close()

    try:
        output_f = zipfile.ZipFile(os.path.join(img_dir, 'output.odp'), 'w')
    except:
        print("Failed to write in " + rep_d + ".")
        return
    for root, dirs, files in os.walk(rep_d):
        for f in files:
            originalPath = os.path.join(root, f)
            zippedPath = os.path.join(root[len(rep_d):], f).lstrip(os.sep)
            try:
                output_f.write(originalPath, zippedPath)
            except:
                print("Failed to write in " + rep_d + ".")
                break
    output_f.close()

    shutil.rmtree(os.path.join(img_dir, '_tempDir'), True)



def parseCommandLine():
    """
    Parses the arguments in the command line in the form:
        dir=imgDir image_size=wxh main_title=title0 sub_title1=title1 sub_title2=title2 first_page_number=page_n

    and returns a dictionary of key-values.
    """
    import sys

    args={}
    for i in range(1, len(sys.argv)):
        values = sys.argv[i].strip().split('=')
        if len(values) == 2:
            args[values[0]] = values[1]

    return args



def get_list(fileName):
    '''
    Returns a list of lines from a text file.
    '''
    result = []
    try:
        f = open(fileName, 'r')
    except:
        return [fileName]
    for line in f:
        f.append(line)
    f.close()
    return result



def main():
    args = parseCommandLine()
    imgDir = args.get("dir")
    size = args.get("image_size")
    position = args.get("image_position")
    title_list0 = get_list(args.get("main_title"))
    title_list1 = get_list(args.get("sub_title1"))
    title_list2 = get_list(args.get("sub_title2"))
    page_n = args.get("first_page_number")

    makeImpress(imgDir, size, position, title_list0, title_list1, title_list2, page_n)



if __name__ == "__main__":
    main()

