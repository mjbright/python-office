
# xls2html.py

xls2html.py is a simple script based on the xlrd module which will create an HTML Table for each Worksheet which exists in a provided Excel Workbook.

## Coloured tables from .xls files
It also attempts to colour cells as per the original Excel Workbook.

**Note**: this can only be done on .xls files as formatting_info is only implemented by xlrd for .xls files.

Colour mapping is not correct, but you can modify the style rules in the resulting HTML to get the correct colours.

If you have an .xlsx file and you wish to convert with colour you will first need to "Save As" in Excel choosing the older "*Excel 97-2003*" .xls format.

**Note**: For informaton about Browser/HTML colour mappings refer to:
- https://www.w3schools.com/colors/colors_names.asp
- https://www.tutorialspoint.com/html/html_colors.htm

## Creating HTML tables from Excel

To convert a Workbook:
    xls2html.py TABLES.xls
or
    xls2html.py TABLES.xlsx

This will produce html files named as per the Worksheet names.

In case of .xls, colour processing will be performed and 2 more HTML files will be produced:
- COLORS.html: lists the colours actually used in the Workbook cells (actual colours may be incorrect)
- COLORS_INDEX.html: list all colours from the Workbooks color_index table

## Example output with the provided TABLES.xls file

### TABLE1:

<table> <tbody> <tr>
  <th class='cellstyle57_9'></th>  <th class='cellstyle57_9'>Heading1</th>  <th class='cellstyle57_9'>Heading2</th>  <th class='cellstyle57_9'>Heading3</th>  <th class='cellstyle57_9'>Heading4</th>
</tr>
<tr>
  <td class='cellstyle57_9'>Row1</td>  <td class='cellstyle64_8'>R1H1</td>  <td class='cellstyle64_8'>R1H2</td>  <td class='cellstyle64_8'>R1H3</td>  <td class='cellstyle64_8'>R1H4</td>
</tr>
<tr>
  <td class='cellstyle57_9'>Row2</td>  <td class='cellstyle64_8'>R2H1</td>  <td class='cellstyle43_8'>R2H2</td>  <td class='cellstyle64_8'>R2H3</td>  <td class='cellstyle64_8'>R2H4</td>
</tr>
<tr>
  <td class='cellstyle57_9'>Row3</td>  <td class='cellstyle64_8'>R3H1</td>  <td class='cellstyle64_8'>R3H2</td>  <td class='cellstyle43_8'>R3H3</td>  <td class='cellstyle64_8'>R3H4</td>
</tr>
<tr>
  <td class='cellstyle57_9'>Row4</td>  <td class='cellstyle64_8'>R4H1</td>  <td class='cellstyle64_8'>R4H2</td>  <td class='cellstyle64_8'>R4H3</td>  <td class='cellstyle64_8'>R4H4</td>
</tr>
<tr>
  <td class='cellstyle57_9'>Row5</td>  <td class='cellstyle64_8'>R5H1</td>  <td class='cellstyle64_8'>R5H2</td>  <td class='cellstyle64_8'>R5H3</td>  <td class='cellstyle64_8'>R5H4</td>
</tr>
 </tbody> </table>


### TABLE2:

<table> <tbody> <tr>
  <th class='cellstyle62_9'>TABLE2</th>  <th class='cellstyle62_9'>Heading1</th>  <th class='cellstyle62_9'>Heading2</th>  <th class='cellstyle62_9'>Heading3</th>  <th class='cellstyle62_9'>Heading4</th>
</tr>
<tr>
  <td class='cellstyle62_9'>Row1</td>  <td class='cellstyle64_8'>R1H1</td>  <td class='cellstyle64_8'>R1H2</td>  <td class='cellstyle64_8'>R1H3</td>  <td class='cellstyle64_8'>R1H4</td>
</tr>
<tr>
  <td class='cellstyle62_9'>Row2</td>  <td class='cellstyle64_8'>R2H1</td>  <td class='cellstyle49_8'>R2H2</td>  <td class='cellstyle64_8'>R2H3</td>  <td class='cellstyle64_8'>R2H4</td>
</tr>
<tr>
  <td class='cellstyle62_9'>Row3</td>  <td class='cellstyle64_8'>R3H1</td>  <td class='cellstyle49_8'>R3H2</td>  <td class='cellstyle49_8'>R3H3</td>  <td class='cellstyle64_8'>R3H4</td>
</tr>
<tr>
  <td class='cellstyle62_9'>Row4</td>  <td class='cellstyle64_8'>R4H1</td>  <td class='cellstyle49_8'>R4H2</td>  <td class='cellstyle64_8'>R4H3</td>  <td class='cellstyle64_8'>R4H4</td>
</tr>
<tr>
  <td class='cellstyle62_9'>Row5</td>  <td class='cellstyle64_8'>R5H1</td>  <td class='cellstyle49_8'>R5H2</td>  <td class='cellstyle64_8'>R5H3</td>  <td class='cellstyle64_8'>R5H4</td>
</tr>
 </tbody> </table>


### COLORS.html:

<table> <tbody> <tr><td> cellstyle57_9: </td><td class='cellstyle57_9'> {  background-color: #339966;  color: #ffffff;} </td></tr>
<tr><td> cellstyle64_8: </td><td class='cellstyle64_8'> {  background-color: #ffffff;  color: #000000;} </td></tr>
<tr><td> cellstyle43_8: </td><td class='cellstyle43_8'> {  background-color: #ffff99;  color: #000000;} </td></tr>
<tr><td> cellstyle62_9: </td><td class='cellstyle62_9'> {  background-color: #333399;  color: #ffffff;} </td></tr>
<tr><td> cellstyle49_8: </td><td class='cellstyle49_8'> {  background-color: #33cccc;  color: #000000;} </td></tr>
 </tbody> </table> 


### COLORS_INDEX.html:

<table> <tbody>
<tr><td> 0: </td><td style='background-color: #000000; color: #ffffff;'> (0, 0, 0) == background-color: #000000; color: #ffffff; </td></tr>
<tr><td> 1: </td><td style='background-color: #ffffff; color: #000000;'> (255, 255, 255) == background-color: #ffffff; color: #000000; </td></tr>
<tr><td> 2: </td><td style='background-color: #ff0000; color: #ffffff;'> (255, 0, 0) == background-color: #ff0000; color: #ffffff; </td></tr>
<tr><td> 3: </td><td style='background-color: #00ff00; color: #ffffff;'> (0, 255, 0) == background-color: #00ff00; color: #ffffff; </td></tr>
<tr><td> 4: </td><td style='background-color: #0000ff; color: #ffffff;'> (0, 0, 255) == background-color: #0000ff; color: #ffffff; </td></tr>
<tr><td> 5: </td><td style='background-color: #ffff00; color: #000000;'> (255, 255, 0) == background-color: #ffff00; color: #000000; </td></tr>
<tr><td> 6: </td><td style='background-color: #ff00ff; color: #000000;'> (255, 0, 255) == background-color: #ff00ff; color: #000000; </td></tr>
<tr><td> 7: </td><td style='background-color: #00ffff; color: #000000;'> (0, 255, 255) == background-color: #00ffff; color: #000000; </td></tr>
<tr><td> 8: </td><td style='background-color: #000000; color: #ffffff;'> (0, 0, 0) == background-color: #000000; color: #ffffff; </td></tr>
<tr><td> 9: </td><td style='background-color: #ffffff; color: #000000;'> (255, 255, 255) == background-color: #ffffff; color: #000000; </td></tr>
<tr><td> 10: </td><td style='background-color: #ff0000; color: #ffffff;'> (255, 0, 0) == background-color: #ff0000; color: #ffffff; </td></tr>
<tr><td> 11: </td><td style='background-color: #00ff00; color: #ffffff;'> (0, 255, 0) == background-color: #00ff00; color: #ffffff; </td></tr>
<tr><td> 12: </td><td style='background-color: #0000ff; color: #ffffff;'> (0, 0, 255) == background-color: #0000ff; color: #ffffff; </td></tr>
<tr><td> 13: </td><td style='background-color: #ffff00; color: #000000;'> (255, 255, 0) == background-color: #ffff00; color: #000000; </td></tr>
<tr><td> 14: </td><td style='background-color: #ff00ff; color: #000000;'> (255, 0, 255) == background-color: #ff00ff; color: #000000; </td></tr>
<tr><td> 15: </td><td style='background-color: #00ffff; color: #000000;'> (0, 255, 255) == background-color: #00ffff; color: #000000; </td></tr>
<tr><td> 16: </td><td style='background-color: #800000; color: #ffffff;'> (128, 0, 0) == background-color: #800000; color: #ffffff; </td></tr>
<tr><td> 17: </td><td style='background-color: #008000; color: #ffffff;'> (0, 128, 0) == background-color: #008000; color: #ffffff; </td></tr>
<tr><td> 18: </td><td style='background-color: #000080; color: #ffffff;'> (0, 0, 128) == background-color: #000080; color: #ffffff; </td></tr>
<tr><td> 19: </td><td style='background-color: #808000; color: #ffffff;'> (128, 128, 0) == background-color: #808000; color: #ffffff; </td></tr>
<tr><td> 20: </td><td style='background-color: #800080; color: #ffffff;'> (128, 0, 128) == background-color: #800080; color: #ffffff; </td></tr>
<tr><td> 21: </td><td style='background-color: #008080; color: #ffffff;'> (0, 128, 128) == background-color: #008080; color: #ffffff; </td></tr>
<tr><td> 22: </td><td style='background-color: #c0c0c0; color: #000000;'> (192, 192, 192) == background-color: #c0c0c0; color: #000000; </td></tr>
<tr><td> 23: </td><td style='background-color: #808080; color: #ffffff;'> (128, 128, 128) == background-color: #808080; color: #ffffff; </td></tr>
<tr><td> 24: </td><td style='background-color: #9999ff; color: #000000;'> (153, 153, 255) == background-color: #9999ff; color: #000000; </td></tr>
<tr><td> 25: </td><td style='background-color: #993366; color: #ffffff;'> (153, 51, 102) == background-color: #993366; color: #ffffff; </td></tr>
<tr><td> 26: </td><td style='background-color: #ffffcc; color: #000000;'> (255, 255, 204) == background-color: #ffffcc; color: #000000; </td></tr>
<tr><td> 27: </td><td style='background-color: #ccffff; color: #000000;'> (204, 255, 255) == background-color: #ccffff; color: #000000; </td></tr>
<tr><td> 28: </td><td style='background-color: #660066; color: #ffffff;'> (102, 0, 102) == background-color: #660066; color: #ffffff; </td></tr>
<tr><td> 29: </td><td style='background-color: #ff8080; color: #000000;'> (255, 128, 128) == background-color: #ff8080; color: #000000; </td></tr>
<tr><td> 30: </td><td style='background-color: #0066cc; color: #ffffff;'> (0, 102, 204) == background-color: #0066cc; color: #ffffff; </td></tr>
<tr><td> 31: </td><td style='background-color: #ccccff; color: #000000;'> (204, 204, 255) == background-color: #ccccff; color: #000000; </td></tr>
<tr><td> 32: </td><td style='background-color: #000080; color: #ffffff;'> (0, 0, 128) == background-color: #000080; color: #ffffff; </td></tr>
<tr><td> 33: </td><td style='background-color: #ff00ff; color: #000000;'> (255, 0, 255) == background-color: #ff00ff; color: #000000; </td></tr>
<tr><td> 34: </td><td style='background-color: #ffff00; color: #000000;'> (255, 255, 0) == background-color: #ffff00; color: #000000; </td></tr>
<tr><td> 35: </td><td style='background-color: #00ffff; color: #000000;'> (0, 255, 255) == background-color: #00ffff; color: #000000; </td></tr>
<tr><td> 36: </td><td style='background-color: #800080; color: #ffffff;'> (128, 0, 128) == background-color: #800080; color: #ffffff; </td></tr>
<tr><td> 37: </td><td style='background-color: #800000; color: #ffffff;'> (128, 0, 0) == background-color: #800000; color: #ffffff; </td></tr>
<tr><td> 38: </td><td style='background-color: #008080; color: #ffffff;'> (0, 128, 128) == background-color: #008080; color: #ffffff; </td></tr>
<tr><td> 39: </td><td style='background-color: #0000ff; color: #ffffff;'> (0, 0, 255) == background-color: #0000ff; color: #ffffff; </td></tr>
<tr><td> 40: </td><td style='background-color: #00ccff; color: #000000;'> (0, 204, 255) == background-color: #00ccff; color: #000000; </td></tr>
<tr><td> 41: </td><td style='background-color: #ccffff; color: #000000;'> (204, 255, 255) == background-color: #ccffff; color: #000000; </td></tr>
<tr><td> 42: </td><td style='background-color: #ccffcc; color: #000000;'> (204, 255, 204) == background-color: #ccffcc; color: #000000; </td></tr>
<tr><td> 43: </td><td style='background-color: #ffff99; color: #000000;'> (255, 255, 153) == background-color: #ffff99; color: #000000; </td></tr>
<tr><td> 44: </td><td style='background-color: #99ccff; color: #000000;'> (153, 204, 255) == background-color: #99ccff; color: #000000; </td></tr>
<tr><td> 45: </td><td style='background-color: #ff99cc; color: #000000;'> (255, 153, 204) == background-color: #ff99cc; color: #000000; </td></tr>
<tr><td> 46: </td><td style='background-color: #cc99ff; color: #000000;'> (204, 153, 255) == background-color: #cc99ff; color: #000000; </td></tr>
<tr><td> 47: </td><td style='background-color: #ffcc99; color: #000000;'> (255, 204, 153) == background-color: #ffcc99; color: #000000; </td></tr>
<tr><td> 48: </td><td style='background-color: #3366ff; color: #ffffff;'> (51, 102, 255) == background-color: #3366ff; color: #ffffff; </td></tr>
<tr><td> 49: </td><td style='background-color: #33cccc; color: #ffffff;'> (51, 204, 204) == background-color: #33cccc; color: #ffffff; </td></tr>
<tr><td> 50: </td><td style='background-color: #99cc00; color: #ffffff;'> (153, 204, 0) == background-color: #99cc00; color: #ffffff; </td></tr>
<tr><td> 51: </td><td style='background-color: #ffcc00; color: #000000;'> (255, 204, 0) == background-color: #ffcc00; color: #000000; </td></tr>
<tr><td> 52: </td><td style='background-color: #ff9900; color: #ffffff;'> (255, 153, 0) == background-color: #ff9900; color: #ffffff; </td></tr>
<tr><td> 53: </td><td style='background-color: #ff6600; color: #ffffff;'> (255, 102, 0) == background-color: #ff6600; color: #ffffff; </td></tr>
<tr><td> 54: </td><td style='background-color: #666699; color: #ffffff;'> (102, 102, 153) == background-color: #666699; color: #ffffff; </td></tr>
<tr><td> 55: </td><td style='background-color: #969696; color: #ffffff;'> (150, 150, 150) == background-color: #969696; color: #ffffff; </td></tr>
<tr><td> 56: </td><td style='background-color: #003366; color: #ffffff;'> (0, 51, 102) == background-color: #003366; color: #ffffff; </td></tr>
<tr><td> 57: </td><td style='background-color: #339966; color: #ffffff;'> (51, 153, 102) == background-color: #339966; color: #ffffff; </td></tr>
<tr><td> 58: </td><td style='background-color: #003300; color: #ffffff;'> (0, 51, 0) == background-color: #003300; color: #ffffff; </td></tr>
<tr><td> 59: </td><td style='background-color: #333300; color: #ffffff;'> (51, 51, 0) == background-color: #333300; color: #ffffff; </td></tr>
<tr><td> 60: </td><td style='background-color: #993300; color: #ffffff;'> (153, 51, 0) == background-color: #993300; color: #ffffff; </td></tr>
<tr><td> 61: </td><td style='background-color: #993366; color: #ffffff;'> (153, 51, 102) == background-color: #993366; color: #ffffff; </td></tr>
<tr><td> 62: </td><td style='background-color: #333399; color: #ffffff;'> (51, 51, 153) == background-color: #333399; color: #ffffff; </td></tr>
<tr><td> 63: </td><td style='background-color: #333333; color: #ffffff;'> (51, 51, 51) == background-color: #333333; color: #ffffff; </td></tr>
<tr><td> 64: </td><td style='background-color: #ffffff; color: #000000;'> None == background-color: #ffffff; color: #000000; </td></tr>
<tr><td> 65: </td><td style='background-color: #ffffff; color: #000000;'> None == background-color: #ffffff; color: #000000; </td></tr>
<tr><td> 81: </td><td style='background-color: #ffffff; color: #000000;'> None == background-color: #ffffff; color: #000000; </td></tr>
<tr><td> 32767: </td><td style='background-color: #ffffff; color: #000000;'> None == background-color: #ffffff; color: #000000; </td></tr>
</tbody> </table>

