#Joseph Widrig and David McPherson
#
#TITLE
#[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')

Function moveZipFile()
{
    
    $location = Get-FileName #-initialDirectory "~\Downloads"
    write $location
    echo "Checking for existance of cxxtest zip file in download"
    if( ($location -ieq "Cancel" ) )
    {
        echo "moving zip file."
        move-item ~\Downloads\cxxtest-4.3.zip ~\cxxt_proj\CxxTest
        return true
    }
    else
    {
        Write-Error "Missing CxxTest Zip file"
        return false
    }
}



#TO ADD: ask whether CxxTest dir exists and then decide whether to do this function
Function createUnzipDirectory()
{
    #$workingDir = "~\cxx_test"
    cd ~
    if(!(Test-Path .\cxx_test -PathType Container) ) #test to see if cxxt_proj already exists
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
    cd ~\cxx_test\CxxTest #move into working directory
    del *.*

}

createUnzipDirectory
if ( moveZipFile )
{ 
    unzip
}
undoCredateUnzipDirectory
#at this point all zip files have been extracted in the CxxTest folder - 10/21/13

Function moveUnzippedFiles()
{
    
}

#~\cxxt_proj\CxxTest\cxxtest-4.3   [\cxxtest] <--- must be moved to C:\Program Files (x86)\MSVS 11.0\VC\include
#                                                                or C:\Program Files\Microsoft SDKs\Windows\v*.*\Include\cxxtest
#  <using microsoft
#  investigate cygwin and mingw
# ADD IN echo(s) OR Write-(s)