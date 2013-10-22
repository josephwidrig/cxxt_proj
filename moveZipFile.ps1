[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')

Function moveZipFile()
{
    Move-Item C:\Users\Joseph\Downloads\cxxtest-4.3.zip C:\Users\Joseph\cxxt_proj\CxxTest
}


#TO ADD: ask whether CxxTest dir exists and then decide whether to do this function
Function createUnzipDirectory()
{
    #move to working directory (can be changed)
    cd C:\Users\Joseph\cxxt_proj
    dir
    mkdir CxxTest
    #[System.IO.Compression.ZipFile]::ExtractToDirectory('~\cxxt_proj\cxxtest-4.3.zip', '~\cxxt_proj\CxxTest')
}



Function maybeUnzip()
{
    ls
    #unzips all compressed files in the directory
    $shell=new-object -com shell.application
    $CurrentLocation=get-location
    #target destination
    cd C:\Users\Joseph\cxxt_proj\CxxTest
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
        $ZipFile.fullname | out-default
        $ZipFolder = $shell.namespace($ZipFile.fullname)
        #error returned from this line
        #(void) changed '$ZipFolder' to '$CurrentLocation'
        $location.Copyhere($ZipFolder.items())
        
    }
}

createUnzipDirectory
moveZipFile
maybeUnzip
#at this point all zip files have been extracted in the CxxTest folder
