#Joseph Widrig and David McPherson
#
#TITLE
#[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')

Function moveZipFile()
{
    #eventually make this version-flexible
    if( (Test-Path ~\Downloads\cxxtest-4.3.zip -PathType leaf) )
    {
    Move-Item ~\Downloads\cxxtest-4.3.zip ~\cxxt_proj\CxxTest
    }
    else
    {
        Write-Error "cxxtest-4.3.zip not found"   
    }
}


#TO ADD: ask whether CxxTest dir exists and then decide whether to do this function
Function createUnzipDirectory()
{
    cd ~
    if(!(Test-Path .\cxxt_proj -PathType Container) ) #test to see if cxxt_proj already exists
    {
        mkdir cxxt_proj
    }
    #move to working directory (can be changed)
    cd ~\cxxt_proj
    if(!(Test-Path .\CxxTest -PathType Container) )
    {
        mkdir CxxTest
    }
    #[System.IO.Compression.ZipFile]::ExtractToDirectory('~\cxxt_proj\cxxtest-4.3.zip', '~\cxxt_proj\CxxTest')
}


#changed from maybeUnzip
Function unzip()
{
    ls
    #unzips all compressed files in the directory
    $shell=new-object -com shell.application #gets a shell object
    $CurrentLocation=get-location
    #target destination
    cd ~\cxxt_proj\CxxTest
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

createUnzipDirectory
moveZipFile
unzip
#at this point all zip files have been extracted in the CxxTest folder - 10/21/13

#~\cxxt_proj\CxxTest\cxxtest-4.3   [\cxxtest] <--- must be moved to C:\Program Files (x86)\MSVS 11.0\VC\include
#                                                                or C:\Program Files\Microsoft SDKs\Windows\v*.*\Include\cxxtest
#  <using microsoft
#  investigate cygwin and mingw
# ADD IN echo(s) OR Write-(s)