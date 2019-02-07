# python-office

Some experiments with Python-Excel and Python-Powerpoint modules

## Python and Excel

There are several Python modules available with different capabilities (different pros and cons).

For information about processing Excel files using Python refer to
- the Python-Excel site: http://www.python-excel.org/
- the Python-Excel Google Group: http://groups.google.com/group/python-excel
- The OpenPyXL Google Group: https://groups.google.com/forum/#!forum/openpyxl-users

**NOTE**: For invoking Python from within Excel you may want to look at [xlwings](https://www.xlwings.org/), also an Open Source project (on [github](https://github.com/ZoomerAnalytics/xlwings).  It runs on Windows or macOS, but not on Linux.

### Python and Excel Examples

- [xlrd-Excel-2-HTMlTable](xlrd-Excel-2-HTMlTable): Example of xlrd use - converts each worksheet into an HTML table

## Python and Powerpoint

The python-pptx module allows to create Powerpoint files from Python.

Chris Moffitt wrote a nice Blog article about ["Creating Powerpoint Presentations with Python" - Aug 2015](http://pbpython.com/creating-powerpoint.html) including some intriguing examples.

It is an active project, for information refer to
- The documentation: https://python-pptx.readthedocs.io/en/latest/
- The PyPi page: https://pypi.org/project/python-pptx/
- The source code on Github: https://github.com/scanny/python-pptx
- Some examples in the documentation QuickStart page: https://python-pptx.readthedocs.io/en/latest/user/quickstart.html

### Python and Powerpoint Examples

- [python-pptx-Quickstart/](python-pptx-Quickstart/): The python-pptx Quickstart examples (almost no changes)

### Using Python under Windows
- [*"Automating PowerPoint With Python"*](http://www.s-anand.net/blog/automating-powerpoint-with-python/)

## PDF output

These libraries do not support saving to PDF format, though there may be ways to perform this task using Windows utlilies.

If you have LibreOffice/OpenOffice installed you may be able to use unoconv.

See:
- [*"StackOverflow discussion"*](https://stackoverflow.com/questions/31487478/how-to-convert-a-pptx-to-pdf-using-python))
- [*"Unoconv repository"*](https://github.com/dagwieers/unoconv)

