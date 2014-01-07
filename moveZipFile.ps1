#Joseph Widrig and David McPherson
#
#TITLE
#[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')

Function moveZipFile()
{
    
    $location = Get-FileName #-initialDirectory "~\Downloads"
    write $location
    echo "Checking for existance of cxxtest zip file in download"
    if( ($location -ne "Cancel" ) )
    {
        echo "moving zip file."
        move-item $location.filename ~\cxx_test\CxxTest
        return $TRUE
    }
    else
    {
        Write-Error "Missing CxxTest Zip file"
        return $FALSE
    }
}



#TO ADD: ask whether CxxTest dir exists and then decide whether to do this function
Function createUnzipDirectory()
{
    #$workingDir = "~\cxx_test"
    cd ~
    if(!(Test-Path .\cxx_test -PathType Container) ) #test to see if cxxt_proj -> cxx_test already exists
    {
        mkdir cxx_test
    }
    #move to working directory (can be changed)
    cd ~\cxx_test
    if(!(Test-Path .\CxxTest -PathType Container) )
    {
        mkdir CxxTest
    }
    #[System.IO.Compression.ZipFile]::ExtractToDirectory('~\cxxt_proj\cxxtest-4.3.zip', '~\cxxt_proj\CxxTest')
}

Function Get-FileName($initialDirectory)
{  
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
 [System.Reflection.Assembly]::LoadWithPartialName("System.Environment") | Out-Null
 
 
 

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.InitialDirectory = $downloads
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $ok = $OpenFileDialog.ShowDialog() #| Out-Null
 if($ok -eq[System.Windows.Forms.DialogResult]::OK)
    {
        $OpenFileDialog.filename
        return $OpenFileDialog
    }
 else
    {
        return "Cancel"
    }
 
 
 #$initialDirectory = $downloads

 #write-output $initialDirectory

 #$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 #$OpenFileDialog.InitialDirectory = $downloads
 #$OpenFileDialog.filter = "All files (*.*)| *.*"
 #$OpenFileDialog.ShowDialog() | Out-Null
 #$OpenFileDialog.filename
} #end function Get-FileName



#changed from maybeUnzip
Function unzip()
{
    ls
    #unzips all compressed files in the directory
    $shell=new-object -com shell.application #gets a shell object
    $CurrentLocation=get-location
    #target destination
    cd ~\cxx_test\CxxTest
    $cxxtestdir=Get-Location
    #changed '=$currentLocation.path' to '$cxxtestdir'
    $CurrentPath=$cxxtestdir.path
    $Location=$shell.namespace($CurrentPath)
    $ZipFiles = get-childitem *.zip
    $ZipFiles.count | out-default
    #define $zipfolder ???
    $ZipFolder=$cxxtestdir
    foreach ($ZipFile in $ZipFiles)
    {
        echo "unzipping:"
        $ZipFile.fullname | out-default
        $ZipFolder = $shell.namespace($ZipFile.fullname)
        #error returned from this line
        #(void) changed '$ZipFolder' to '$CurrentLocation'
        $location.Copyhere($ZipFolder.items())
        
    }
    echo "The cxxtest files have been extracted to cxxtest 4.3"
}

Function undoCreateUnzipDirectory
{
    cd ~\cxx_test #move into working directory
    del -r *
    cd ..
    Remove-item cxx_test

}

Function moveUnzippedFiles()
{
    $Stem = 'C:\Program Files (x86)\'
    $Directories = Get-ChildItem $Stem -Filter 'Microsoft Visual Studio*' -Name
    $DirectoryCount= (Get-ChildItem $Stem -Filter 'Microsoft Visual Studio*').Count
    Sort-Object $Directories
    echo $Directories
    $cxxdir = '\VC\include\cxxtest'
    $fullpath = $Directories[0]
    if( $DirectoryCount -gt 1 ) #prompt user
    {
        echo "Multiple MSVS Versions detected."
        
        echo $fullpath
        $i = 1
        foreach( $Directory in $Directories )
        {
            Write $i": "$Directory
            $i++
        }
        $which = $DirectoryCount+1
        while (!($which-1 -lt $DirectoryCount -and $which-1 -ge 0))
        {
        $which = Read-Host "Enter the number corresponding to the version you would like to use."
        }
        $fullpath = $directories[$which-1]
    }
    $fullpath = $Stem + $fullpath + $cxxdir
    echo $fullpath

    #change to reflect choice by $fullpath
    if(!(Test-Path $fullpath -PathType Container) ) #test to see if 'C:\Program Files (x86)\Microsoft Visual Studio 11.0\VC\include\cxxtest' already exists
    {
        echo "cxxtest Directory not found in Visual Studio, making directory"
        New-Item $fullpath -ItemType Directory
    }
    echo "Copying items to new directory"
    Copy-Item cxxtest-4.3\cxxtest\* 'C:\Program Files (x86)\Microsoft Visual Studio 11.0\VC\include\cxxtest'
    if(!(Test-Path ~\cxxtest -PathType Container) ) #test to see if 'C:\Program Files (x86)\Microsoft Visual Studio 11.0\VC\include\cxxtest' already exists
    {
        New-Item ~\cxxtest -ItemType Directory
    }
    Copy-Item cxxtest-4.3 ~\cxxtest -Recurse
}

#MAIN
createUnzipDirectory
if ( moveZipFile )
{ 
    unzip
    moveUnzippedFiles
    #move files to correct location
         #include dir to include location
         #rest of files to home directory
}
undoCreateUnzipDirectory
#at this point all zip files have been extracted in the CxxTest folder - 10/21/13



#~\cxxt_proj\CxxTest\cxxtest-4.3   [\cxxtest] <--- must be moved to C:\Program Files (x86)\MSVS 11.0\VC\include
#                                                                or C:\Program Files\Microsoft SDKs\Windows\v*.*\Include\cxxtest
#  <using microsoft
#  investigate cygwin and mingw
# ADD IN echo(s) OR Write-(s)

# 1/7/14 - This program must be run in a 'run as administrator' powershell window.