#!/usr/bin/env python3

import datetime, xlrd
import sys

VERBOSE=False

HTML_PAGE=False
HTML_PAGE=True

PROCESS_COLOR=True

if len(sys.argv) > 1 and sys.argv[1] == '-v':
    VERBOSE=True

def die(msg):
    print("die: " + msg)
    sys.exit(1)

COLOR_TABLE={}

TABLE_HEADER=''' <table> <tbody> '''

TABLE_FOOTER=''' </tbody> </table> '''

'''
    convertDate(isoDate): Concert string in iso format YYYY-MM-DD to DD-Mon-YYYY
'''
def convertDate(isoDate):
    bits=isoDate.split("-")
    mth=int(bits[1])-1
    if mth < 0 or mth > 11:
        die("Failed to parse month[{}] in date[{}]".format(bits[1], isoDate))

    month=['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][mth]

    return bits[2]+" "+month+" "+bits[0]

'''
    csv_from_excel(wb, worksheet, txt_opfile): Export Worksheet as a CSV file
'''
def csv_from_excel(wb, worksheet, txt_opfile):

    sheet = wb.sheet_by_name(worksheet)
    op = open(txt_opfile, "w")

    #print("Sheet: " + worksheet)
    rownum=0
    op.write(", ".join(sheet.row_values(rownum)))

    for rownum in range(1,sheet.nrows):
        op.write(", ".join(sheet.row_values(rownum)))

    op.close()

'''
    Calculate brightness as %age based on RGB tuple (r,g,b)
'''
def brightnessPC(fgbg, tple):
    if tple == None:
        return 0.0

    #print("{}: TPLE={}".format(fgbg, tple))
    brightness=100.0*( (tple[0]*tple[0]) + (tple[1]*tple[1]) + (tple[2]*tple[2]) ) / (3.0 * 255 * 255)

    #print("brightness=" + str(brightness))
    return brightness


'''
    Convert RGB tuple (r,g,b) to hex string: e.g. #ffffff for white
'''
def To0xRGB(tple):
    if tple == None:
        return ''

    return "#%02x%02x%02x" % tple

'''
    dump_style_rules(color_dict): Dump color table as <STYLE> ... </STYLE> section
'''
def dump_style_rules(color_dict, default_color='#ffffff'):
    style='<STYLE>\n'

    for idx in color_dict:
        color = color_dict[idx]
        if color == None or color == '':
            color=default_color

        style += '.cellstyle{} {}\n'.format(idx, color)

    style += '</STYLE>\n'
    return style

'''
    getCellColors(wb, sheet_name, sheet, row, col): Get foreground/background colours for Cell:
'''
def getCellColors(wb, sheet_name, sheet, row, col):
    xfx = sheet.cell_xf_index(row, col)
    xf  = wb.xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    #print("bgx=" + str(bgx))
    #bgx = xf.background.background_colour_index

    font = wb.font_list[xf.font_index]
    #font.dump()
    fgx  = font.colour_index
    #print("fgx=" + str(fgx))

    pos="(%s,%s)" % (row,col)
    background = ""
    obg = ""
    cell_info=""
    SHOW_ALL=False

    if bgx in wb.colour_map:
        background_tpl = wb.colour_map[bgx]
        background = To0xRGB(background_tpl)
    else:
        print("colour_map keys=<{}>".format(join(",", wb.colour_map.keys())))
        die("Not found bgx={}".format(bgx))

    font_color = ""
    ofg = ""

    if fgx in wb.colour_map:
        font_color_tpl = wb.colour_map[fgx]
        font_color = To0xRGB(font_color_tpl)
        color  = wb.colour_map.get(fgx)
    else:
        print("colour_map keys=<{}>".format(join(",", wb.colour_map.keys())))
        die("Not found fgx={}".format(fgx))

    style_index = "{}_{}".format(bgx, fgx)

    if background == '' and font_color == '':
        background = '#ffffff'
        font_color = '#000000'

    if background == '':
        if brightnessPC('font', font_color_tpl) < 50.0:
            background = '#ffffff'
        else:
            background = '#000000'

    if font_color == '':
        if brightnessPC('background', background_tpl) < 50:
            font_color = '#000000'
        else:
            font_color = '#ffffff'

    print("style_index={}: background={} font_color={}".format(style_index, background,font_color))
    json_style='{\n  ' + 'background-color: {};\n  color: {};\n'.format(background,font_color) + '}'

    if style_index in COLOR_TABLE:
        if COLOR_TABLE[style_index] != json_style:
            die("BUG-json_style")
    else:
        COLOR_TABLE[style_index]   = json_style

    style = " class='cellstyle{}'".format(style_index)

    if cell_info != "":
        cell_info = sheet_name + ": " + pos + " " + cell_info
        print(cell_info)
        #xf.dump()
        #font.dump()

    return style

'''
    def html_from_excel(wb, sheet_name, html_opfile, HTML_PAGE=False):
        Write out Excel table spreadsheet as an HTML table, with coloured cells
'''
def html_from_excel(wb, sheet_name, html_opfile, HTML_PAGE=False):

    sheet = wb.sheet_by_name(sheet_name)
    op = open(html_opfile, "w")

    table_html = "\n<h1>Sheet: " + sheet_name + "</h1>"
    table_html += TABLE_HEADER

    rownum=0
    table_html += "<tr>\n  "

    for col in range( len( sheet.row_values(rownum) ) ):
        if col > 0:
            table_html += "</th>  "

        if PROCESS_COLOR:
            table_html += "<th{}>".format( getCellColors(wb, sheet_name, sheet, rownum, col) )
        else:
            table_html += "<th>"

        table_html += sheet.row_values(rownum)[col]

    table_html += "</th>\n</tr>\n"

    for rownum in range(1,sheet.nrows):
        table_html += "<tr>\n  "

        for col in range( len( sheet.row_values(rownum) ) ):
            if col > 0:
                table_html += "</td>  "

            if PROCESS_COLOR:
                table_html += "<td{}>".format( getCellColors(wb, sheet_name, sheet, rownum, col) )
            else:
                table_html += "<td>"

            table_html += sheet.row_values(rownum)[col]
        table_html += "</td>\n</tr>\n"

    table_html += TABLE_FOOTER

    HEADER_HTML='!<DOCTYPE HTML>\n<HTML>\n'
    page_html = HEADER_HTML

    if PROCESS_COLOR:
        page_html += dump_style_rules( COLOR_TABLE )
    page_html += "<BODY>"
    page_html += table_html
    page_html += "</BODY>\n</HTML>"

    if HTML_PAGE:
        op.write(page_html)
    else:
        op.write(table_html)

    op.close()

'''
    write_color_index_table(wb, html_opfile):
            Write out Workbook colour_map as a coloured HTML table
 '''
def write_color_index_table(wb, html_opfile):

    op = open(html_opfile, "w")

    op.write("<HTML>\n<BODY>")
    op.write(TABLE_HEADER)

    for idx in wb.colour_map:
        background_tpl = wb.colour_map[idx]
        if background_tpl:
            background_tpl = wb.colour_map[idx]
            background = To0xRGB( background_tpl )

            if brightnessPC('background', background_tpl) < 50.0:
                font_color = '#ffffff'
            else:
                font_color = '#000000'
        else:
            #print("tpl={} bg={}".format( background_tpl, background))
            background = '#ffffff'
            font_color = '#000000'

        style='background-color: {}; color: {};'.format(background, font_color)
        op.write("<tr><td> {}: </td><td style='{}'> {} == {} </td></tr>\n".format(idx, style, background_tpl, style))

    op.write(TABLE_FOOTER)
    op.write("</BODY>\n</HTML>")
    op.close()


'''
    write_color_table(wb, html_opfile):
        Write out COLOR_TABLE as a coloured HTML table
'''
def write_color_table(wb, html_opfile):

    op = open(html_opfile, "w")

    op.write("<HTML>\n" + dump_style_rules( COLOR_TABLE ) + "\n<BODY>")

    op.write(TABLE_HEADER)

    #for idx in wb.colour_map:
    for style_index in COLOR_TABLE:
        style = COLOR_TABLE[style_index].replace("\n","")
        #op.write("<tr><td> {}: </td><td style='{}'> {} SOME TEXT </td></tr>\n".format(style_index, style, style))
        op.write("<tr><td> cellstyle{}: </td><td class='cellstyle{}'> {} </td></tr>\n".format(style_index, style_index, style))

    op.write(TABLE_FOOTER)

    op.write("</BODY>\n</HTML>")
    op.close()

################################################################################
# Main:

xls_file = sys.argv[1]

if xls_file == "-table":
    HTML_PAGE=False
    xls_file = sys.argv[2]

if xls_file == "-page":
    HTML_PAGE=True
    xls_file = sys.argv[2]


'''
    formatting_info option only works on .xls not on .xlsx:

    NOTE: colouring will not work without formatting_info, so not on .xlsx files
'''

if ".xlsx" in xls_file:
    PROCESS_COLOR=False
    wb = xlrd.open_workbook(xls_file)
else:
    wb = xlrd.open_workbook(xls_file, formatting_info=True)

sheet_names = wb.sheet_names()
print(sheet_names)

for sheet_name in sheet_names:
    html_from_excel(wb, sheet_name, sheet_name+".html", HTML_PAGE)
    #csv_from_excel(wb, sheet_name, sheet_name+".csv")

if PROCESS_COLOR:
    write_color_index_table(wb, "COLORS_INDEX.html")
    write_color_table(wb, "COLORS.html")

