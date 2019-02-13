# TO ADD

## Python-Excel

- Examples of Excel-2-HTML conversion for different Python-Excel libraries
  - highlight strengths/weaknesses
  
- Examples of use of different Python-Excel libraries to copy/create Excel files
  - highlight strengths/weaknesses
    - openpyxl: can read/write .xlx/.xlsx files (but buggy)
    - xlrd: can copy worksheets, but copies values only not formulas, not formats

## Python-Powerpoint

- Python-PPTX examples
  - More elaborate slide creaton examples
  - Read existing .PPTX, attempt French-English translation
  - Read existing .PPTX, attempt conversion to javascript slideset, e.g. Remark.js, WebSlides
  - Read existing javascript slideset, e.g. Remark.js, WebSlides, generate .PPTX
  
<hr/>

## Experiments

### Embedding SVG - not allowed (by github.com sanitizer):

<h1>Example of embedded SVG</h1> <table> <tbody> <tr>
  <th></th>  <th></th>  <th></th> 
</tr>
<tr>
  <td>img tag: (png)<img width=100 src='https://www.shareicon.net/download/2015/05/18/40183_text_512x512.png'/></td>
  <td>img tag: (svg)<img width=100 src='https://upload.wikimedia.org/wikipedia/commons/thumb/0/02/SVG_logo.svg/330px-SVG_logo.svg.png'/></td>
  <td>
      <svg height="30" width="200" xmlns:xlink="http://www.w3.org/1999/xlink">
	    <a xlink:href="https://www.w3schools.com/graphics/" target="_blank">
	    <text x="0" y="15" fill="red" text-decoration="underline"> Embedded SVG is not rendered  - only element text (a link) is shown on github </text>
	    </a>
      </svg>
  </td>
</tr>
</tbody> </table>

### Linking to SVG - not allowed (by github.com sanitizer):

<h1>Example of embedded SVG</h1> <table> <tbody> <tr>
  <th></th>  <th></th>  <th></th> 
</tr>
<tr>
  <td>img tag:<img width=100 src='https://www.shareicon.net/download/2015/05/18/40183_text_512x512.png'/></td>
  <td>
	  <img src='images/memDistrib-jemalloc-4.0.3.svg'/>
	</td> <td>
	  <svg height="30" width="200" xmlns:xlink="http://www.w3.org/1999/xlink">
	      <a xlink:href="https://www.w3schools.com/graphics/" target="_blank">
		<text x="0" y="15" fill="red" text-decoration="underline"> Embedded SVG is not rendered  - only element text (a link) is shown on github </text>
				  </a>
	  </svg>
  </td>
</tr>
<tr>
  <td><img src='./images/test.svg'/></td>
  <td><img src='./images/test.svg'/></td>
  <td><img src='./images/test.svg'/></td>
</tr>
 </tbody> </table>


