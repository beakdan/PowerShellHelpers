# PowerShellHelpers
A simple attempt to put order to the collection of scripts that I use on a daily basis.

## Export-ExcelXLSX
I really like how simple is to create Worksheets usign [ClosedXML](https://github.com/ClosedXML/ClosedXML) assembly. So I wrote this little function to produce xlsx files from the pipeline.
ClosedXML can add a DataTable as a Worksheet and I heavily use that feature to dump complete resultsets, so this function also accepts an array of DataTables.
