# plotExcel 1.0.0

* Add HTML plot support by rendering `.html` and `.htm` files to PDF with `pagedown::chrome_print()`.
* Rename the package from `excelPlot` to `plotExcel`.
* Rename `diffpdf()` to `plotExcelDiff()` and remove the pre-1.0 name.
* Keep the public API focused by making low-level file and page helpers internal.
* Use portable temporary output paths and consistent supported-file messaging.
* Document plot, text, and diff decorator syntax in `availableStyles()`.

# excelPlot 0.2.3

* Add support for DOCX, PPTX, and Excel plot inputs in the preprocessing pipeline by converting to PDF first via a new `pngPipelineConvertOfficeToPdf` stage.
* Conversion backends:
    * Windows: Office COM automation via PowerShell (`Word.Application`, `PowerPoint.Application`, and `Excel.Application`)
    * Linux/macOS: headless LibreOffice (`libreoffice --headless --convert-to pdf`)
* Add `convertOfficeToPdf()` with `single` and `A4` Excel pagination modes.
* Make `plotExcel(FLAGpdf = TRUE)` cross-platform.
* `plotExcelFolder()` now discovers `docx` and `pptx` files in addition to `pdf` and `png`.

# excelPlot 0.2.2

* diffpdf as a direct UI function for diffing two files.
* Catch case of path = "." in plotExcelFolder
* Add FLAGopenExcel and FLAGtemp to plotExcel directly.

# excelPlot 0.2.1

* Add diffing capability. plotExcelFolder now also has compareToCommit argument to compare diffing natively.

# excelPlot 0.2.0

* Add plotExcelFolder and compareProjects_excelPlot

# excelPlot 0.1.9

* Catching missing files by inserting a dummy picture.

# excelPlot 0.1.0

* First version of the package.
* Plot preprocessing pipeline supports:
    * Checking out from git
    * Extracting pages from dpf
    * Cropping
    * All functions are idempotent - if an output file exists already for a given unchanged input, nothing is done. This will make rerunning the function very fast.
* Text styles are available:
    * bold and large: Left, center, rotate up, rotate down
    * plain
