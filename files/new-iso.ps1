<#
.SYNOPSIS
  Creates a new iso file on disk.

.DESCRIPTION
  This cmdlet creates a new .iso file using .NET public class ISOFILE.

.PARAMETER Source
  One or more file references that are to be added to the .ISO file.
  FileInfo objects (as generated by Get-Item etc) or strings can be passed
  in via the pipeline to be added to the iso file.

.PARAMETER Path
  Optional. The output file path that the iso file will be saved in. It is also
  used for resolving any ambiguous file references, i.e. any file passed in
  via file name and not full path.

  If not specified the current working directory is used for the output file
  and attempting to resolve all ambiguous file references.

.PARAMETER Name
  Optional. The name of the file generated. If no name is provided the current timestamp will be substituted. 

.PARAMETER Clipboard
  Optional. Boolean variable that using a different parameter set.

.EXAMPLE
  .\new-iso.ps1 -Source c:\Windows\Temp
  This command creates a .iso file in $env:temp folder (default location) that contains c:\tools and c:\downloads\utils folders.
  The folders themselves are included at the root of the .iso image.  

.EXAMPLE
  dir c:\WinPE | .\new-iso.ps1 -Path c:\temp\WinPE.iso -BootFile "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\efisys.bin" -Media DVDPLUSR -Title "WinPE" 
  This command creates a bootable .iso file containing the content from c:\WinPE folder, but the folder itself isn't included. Boot file etfsboot.com can be found in Windows ADK. Refer to IMAPI_MEDIA_PHYSICAL_TYPE enumeration for possible media types: http://msdn.microsoft.com/en-us/library/windows/desktop/aa366217(v=vs.85).aspx

.LINK
  https://gallery.technet.microsoft.com/scriptcenter/New-ISOFile-function-a8deeffd 
#>

[CmdletBinding(DefaultParameterSetName = 'Source')]
  Param( 
    [Parameter(HelpMessage = "Items to include in iso file.", Position = 0, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Source')]
    [ValidateNotNullOrEmpty()]
    [string] $Source,

    [Parameter(HelpMessage = "Directory to put the file.", Position = 2, Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Source')]
    [string] $Destination = $(Get-Location),

    [Parameter(HelpMessage = "Name of resulting file.", Position = 1, Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Source')]
    [string] $Name = $(Get-Date).ToString("yyyyMMdd-HHmmss.ffff") + ".iso",

    [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})][string]$BootFile = $null, 
    [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER', 
    [string] $Title = $(Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
    [switch] $Force, 
    [parameter(ParameterSetName='Clipboard')]
    [switch] $FromClipboard 
  ) 
function make
{ 
  Start-Transcript -Path (Get-Location)
  Begin 
  {  
    ($cp = new-object System.CodeDom.Compiler.CompilerParameters).CompilerOptions = '/unsafe' 
    if (!('ISOFile' -as [type])) {  
      Add-Type -CompilerParameters $cp -TypeDefinition @' 
public class ISOFile  
{ 
  public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks)  
  {  
    int bytes = 0;  
    byte[] buf = new byte[BlockSize];  
    var ptr = (System.IntPtr)(&bytes);  
    var o = System.IO.File.OpenWrite(Path);  
    var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;  
  
    if (o != null) { 
      while (TotalBlocks-- > 0) {  
        i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);  
      }  
      o.Flush(); o.Close();  
    } 
  } 
}  
'@  
    } 
  
    if ($BootFile) { 
      if('BDR','BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
      ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type=1}).Open()  # adFileTypeBinary 
      $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
      ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
    } 
 
    $MediaType = @('UNKNOWN','CDROM','CDR','CDRW','DVDROM','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','HDDVDROM','HDDVDR','HDDVDRAM','BDROM','BDR','BDRE') 
 
    Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))" 
    ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName=$Title}).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 
    
    if (!($Target = New-Item -Path $(Join-Path -Path $Destination -ChildPath $Name) -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break } 
  }  
 
  Process 
  { 
    Try 
    {
      if($FromClipboard) 
      { 
        if($PSVersionTable.PSVersion.Major -lt 5) 
        { 
          Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break 
        } 
        $Source = Get-Clipboard -Format FileDropList 
      } 

      foreach ($item in Get-ChildItem -Path $Source) 
      { 
        if($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) 
        { $item = Get-Item -LiteralPath $item } 

        if($item) 
        { 
          Write-Verbose -Message "Adding item to the target image: $($item.FullName)" 
          try { $Image.Root.AddTree($item.FullName, $true) } 
          catch { Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') } 
        } 
      } 
    }
    Catch
    {

    }
    Finally
    {
      Stop-Transcript 
    }
    
  } 
 
  End 
  {  
    if ($Boot) { $Image.BootImageOptions=$Boot }  
    $Result = $Image.CreateResultImage()  
    [ISOFile]::Create($Target.FullName,$Result.ImageStream,$Result.BlockSize,$Result.TotalBlocks) 
    Write-Verbose -Message "Target image ($($Target.FullName)) has been created" 
    $Target
    return 0
    Exit
  } 
}
make 