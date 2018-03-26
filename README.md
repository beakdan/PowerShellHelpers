# PowerShellHelpers
A simple attempt to put order to the collection of scripts that I use on a daily basis.

## Export-ExcelXLSX.ps1
I really like how simple is to create Worksheets usign [ClosedXML](https://github.com/ClosedXML/ClosedXML) assembly. So I wrote this little function to produce xlsx files from the pipeline.
ClosedXML can add a DataTable as a Worksheet and I heavily use that feature to dump complete resultsets, so this function also accepts an array of DataTables.

## Git.ps1
Some helper functions for common task executed on git repositories. Although function names were defined to honor the aproved PS verb list, I defined alias that are more clear.

* Find-GitCurrentBranch (Alias: *GitCurrentBranch*)
: Returns the name of the branch currently checked out

* Request-GitBranch (Alias: *GitCheckout*)
: Checks out the specified branch if it's not already checked out and returns an informative message

* Sync-GitLocalRepository (Alias: *GitFetch*)
: Executes a git Fetch and returns an informative status string (*behind*, *up-to-date* or *ahead*)

* Merge-GitWorkspaceFromRemote (Alias: *GitPull*)
: Executes a Pull and returns the command output

* Get-GitFileLastCommit (Alias: *GitFileLastCommit*)
: Gets the last commit date from a file at UTC 0

