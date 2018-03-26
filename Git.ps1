#Author: Adán Bucio
#Dependencies: Don't forget to install git first

$gitExe = "$env:ProgramFiles\Git\bin\git.exe";

function Find-GitCurrentBranch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
            [string]$RepoPath
	)

    $gitParams = @('-C', """$($RepoPath.TrimEnd("\"))""", 'symbolic-ref', '--short', 'HEAD');
    return [string](& $gitExe $gitParams);
}

#Git Checkout
function Request-GitBranch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
            [string]$RepoPath,
        [Parameter(Mandatory=$true)]
            [string]$Branch
	)

    if(-not [System.IO.Directory]::Exists($RepoPath)) {
        throw New-Object System.IO.DirectoryNotFoundException -ArgumentList $RepoPath
    }

    $gitParams = @('-C', """$($RepoPath.TrimEnd("\"))""", '-c', 'diff.mnemonicprefix=false', '-c', 'core.quotepath=false');

    $status = "Already in branch '$Branch'";

    if($Branch -ne (Find-GitCurrentBranch -RepoPath $RepoPath)){
        $status = [string](
            & $gitExe $gitParams checkout $Branch 2>&1 | % {
                if($_ -is [System.Management.Automation.ErrorRecord]) {
                    #Ps sees Branch checkout as an exception
                    $_.Exception.Message
                }
            });
        & $gitExe $gitParams submodule update --init --recursive *>$null
    }

    return $status;
}

#Git Fetch
function Sync-GitLocalRepository {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
            [string]$RepoPath
	)

    $gitParams = @('-C', """$($RepoPath.TrimEnd("\"))""", '-c', 'diff.mnemonicprefix=false', '-c', 'core.quotepath=false');

    & $gitExe $gitParams fetch origin *>$null;
    $StatusMsg = [string](& $gitExe $gitParams status 2>$null);
    $BranchStatusMatch = [regex]::Match($StatusMsg, '(?:branch is )(?<status>[^\s]+) ', 
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor 
            [System.Text.RegularExpressions.RegexOptions]::Multiline);

    return $BranchStatusMatch.Groups[1].Value;
}

#Git Pull
function Merge-GitWorkspaceFromRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
            [string]$RepoPath,
        [Parameter(Mandatory=$true, ParameterSetName='Current')]
            [switch]$UseCurrentBranch,
        [Parameter(Mandatory=$true, ParameterSetName='Specify')]
            [string]$Branch
	)

    if($UseCurrent.IsPresent) {
        $Branch = (Find-GitCurrentBranch -RepoPath $RepoPath);
    }

    $gitParams = @('-C', """$($RepoPath.TrimEnd("\"))""", '-c', 'diff.mnemonicprefix=false',
        '-c', 'core.quotepath=false', 'pull', 'origin', $Branch);

    return [string](& $gitExe $gitParams)
}

function Get-GitFileLastCommit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
            [string]$RepoPath,
        [Parameter(Mandatory=$true)]
            [string]$FileRelativePath
	)

    $gitParams = @('-C', """$($RepoPath.TrimEnd("\"))""", 'log', '-1', '--format="%aI"', '--', """$FileRelativePath""");

    $file = [System.IO.Path]::Combine($RepoPath, $FileRelativePath);
    if(-not [System.IO.File]::Exists($file)) {
        throw New-Object System.IO.FileNotFoundException `
            -ArgumentList ([System.IO.Path]::GetFileName($file));
    }

    $gitOutput = [string](& $gitExe $gitParams);
    $result = if([string]::IsNullOrWhiteSpace($gitOutput)) { [datetime]::MinValue } else {
                    [System.DateTimeOffset]::ParseExact($gitOutput, 'yyyy-MM-ddTHH:mm:sszzz',
                        [cultureinfo]::InvariantCulture
                    ).UtcDateTime
                };
    return $result
}

<#
Functions defined with standard verbs but, it's easier
to use the functions aliased with Git* prefix
#>
Set-Alias -Name GitCurrentBranch -Value Find-GitCurrentBranch
Set-Alias -Name GitCheckout -Value Request-GitBranch
Set-Alias -Name GitFetch -Value Sync-GitLocalRepository
Set-Alias -Name GitPull -Value Merge-GitWorkspaceFromRemote
Set-Alias -Name GitFileLastCommit -Value Get-GitFileLastCommit

#Export-ModuleMember -Function Find-GitCurrentBranch, Request-GitBranch, Sync-GitLocalRepository,
#                        Merge-GitWorkspaceFromRemote, Get-GitFileLastCommit `
#                    -Alias GitCurrentBranch, GitCheckout, GitFetch, GitPull, GitFileLastCommit;

<#
#Usage:

GitCurrentBranch -RepoPath 'E:\Repositories\MyRepo'

GitCheckout -RepoPath 'E:\Repositories\MyRepo' -Branch master

GitFetch -RepoPath 'E:\Repositories\MyRepo'

GitPull -RepoPath 'E:\Repositories\MyRepo' -UseCurrentBranch
GitPull -RepoPath 'E:\Repositories\MyRepo' -Branch myBranch

GitFileLastCommit -RepoPath 'E:\Repositories\MyRepo' -FileRelativePath 'folder\myfile.xxx'
#>