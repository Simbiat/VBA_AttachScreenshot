# VBA_AttachScreenshot
Small VBA code sample to attach screenshot into an Excel Cell

Just paste the code into a module, set the cell you want to use (adjust `CellToPaste`) and that's it. Calling the `AttachScreenshot`, will paste image from your buffer, creatae temporary chart area of appropriate size as the screenshot, paste the image onto chart, export chart to a temporary file and then attach it as object to the cell and remove all the temporary objects.

Sadly, this code will not work on Shared Workbooks, since you can't insert objects there.
