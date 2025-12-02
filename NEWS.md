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
