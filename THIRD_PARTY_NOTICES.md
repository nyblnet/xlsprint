# Third-party notices

## Microsoft MicroTimer sample

`xlsprint/vba/XLSprintTimer.bas` contains the `MicroTimer` function, adapted
from Microsoft's article "Excel performance: Improving calculation
performance":
<https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance>.
Code samples in the MicrosoftDocs/VBA-Docs repository are licensed under the
MIT License:

    The MIT License (MIT)
    Copyright (c) Microsoft Corporation

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to
    deal in the Software without restriction, including without limitation the
    rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
    sell copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
    IN THE SOFTWARE.

## Runtime dependencies

| Package | License | Use |
|---|---|---|
| openpyxl | MIT | Static workbook inspection |
| et_xmlfile | MIT | openpyxl dependency |
| pywin32 (optional, Windows) | PSF-2.0 | COM automation of Excel |
| pytest (dev only) | MIT | Tests |

XLSprint is not affiliated with FastExcel or Decision Models Ltd and contains
none of their code. FastExcel is referenced only as a point of comparison for
purpose.
