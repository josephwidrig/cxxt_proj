#Joseph Widrig and David McPherson
#
#TITLE


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


#Purpose:
#tests to see if working directory exists. IF doesn't then makes it. Moves into working directory, creates (if doesn't exist) an extraction directory.
Function createUnzipDirectory()
{
    #$workingDir = "~\cxx_test"
    echo "Moving to home directory"
    cd ~
    echo "Checking for working directory existence \'\\cxx_test\'"
    if(!(Test-Path .\cxx_test -PathType Container) ) #test to see if cxxt_proj -> cxx_test already exists
    {
        echo "Directory not found."
        echo "Creating working direcotry '~\cxx_test'"
        mkdir cxx_test
    }
    #move to working directory (can be changed)
    echo "Moving to working direcotry"
    cd ~\cxx_test
    echo "Checking for extraction directory '\CxxTest'"
    if(!(Test-Path .\CxxTest -PathType Container) )
    {
        echo "Extraction directory not found. Creating directory '~\cxx_test\CxxTest'"
        mkdir CxxTest
    }
   
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
}



#changed from maybeUnzip
Function unzip()
{

    #unzips all compressed files in the directory
    $shell=new-object -com shell.application #gets a shell object
    #target destination
    cd ~\cxx_test\CxxTest
    $cxxtestdir=Get-Location
    $CurrentPath=$cxxtestdir.path
    #getting namespace object for later copying
    $Location=$shell.namespace($CurrentPath)
    $ZipFiles = get-childitem *.zip
    $ZipFiles.count | out-default
    foreach ($ZipFile in $ZipFiles)
    {
        echo "unzipping:"
        $ZipFile.fullname | out-default
        $ZipFolder = $shell.namespace($ZipFile.fullname)
        #error returned from this line
        
        $location.Copyhere($ZipFolder.items())
        
    }
    echo "The cxxtest files have been extracted"
}

Function undoCreateUnzipDirectory
{
    echo "removing temporary files"
    cd ~\cxx_test #move into working directory
    del -r *
    cd ..
    echo "removing temporary directories"
    Remove-item cxx_test

}

Function moveUnzippedFiles()
{
    $Stem = 'C:\Program Files (x86)\'
    $cxxdir = '\VC\include\cxxtest'
    $Directories = Get-ChildItem $Stem -Filter 'Microsoft Visual Studio*' -Name
    $DirectoryCount= (Get-ChildItem $Stem -Filter 'Microsoft Visual Studio*').Count
    Sort-Object $Directories
    $fullpath = $Directories[0] #if only 1 MSVS version installed, use that one
    if( $DirectoryCount -gt 1 ) #else prompt user to choose which one
    {
        echo "Multiple MSVS Versions detected."
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
    

    #change to reflect choice by $fullpath
    if(!(Test-Path $fullpath -PathType Container) ) #test to see if 'C:\Program Files (x86)\Microsoft Visual Studio 11.0\VC\include\cxxtest' already exists
    {
        echo "CxxTest Directory not found in Visual Studio, making directory"
        New-Item $fullpath -ItemType Directory
    }
    echo "Copying CxxTest to MSVS include directory"
    Copy-Item cxxtest-4.3\cxxtest\* $fullpath
    echo "Copying CxxTest files to home directory"
    if(!(Test-Path ~\cxxtest -PathType Container) ) #test to see if 'C:\Program Files (x86)\Microsoft Visual Studio 11.0\VC\include\cxxtest' already exists
    {
        New-Item ~\cxxtest -ItemType Directory
    }
    Copy-Item cxxtest-4.3 ~\cxxtest -Recurse -Force 
}

#MAIN
createUnzipDirectory
if ( moveZipFile )
{ 
    unzip
    moveUnzippedFiles
 
}
undoCreateUnzipDirectory
# 1/7/14 - This program must be run in a 'run as administrator' powershell window.