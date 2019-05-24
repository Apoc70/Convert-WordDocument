  <#
    .Synopsis
    PowerShell script to convert Word documents

    .Description
    Convert a collection of Word compatible documents to a selected format, using SaveAs in Word.

    Currently supported target document types:
    - Default --> Word 2016
    - PDF
    - XPS
    - HTML

    Version 1.0 2017-08-31 - Thomas Stensitzki - Initial Release
    PROPOSED Version 2.0 2019-05-24 - Ali Robertson's Fork 
    
    .NOTES 
  
    Requirements 
    - Word 2016+ installed locally

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0      Initial release
    2.0      Copy file attributes from old to new file. Add InstancePerFile, ContinueOnError, and fix DeleteExistingFiles. 

    .LINK
    https://www.granikos.eu/en/justcantgetenough/PostId/353/convert-word-documents-using-powershell

    .PARAMETER SourcePath
    Source path to a folder containing the documents to convert or full path to a single document
    
    .PARAMETER IncludeFilter
    File extension filter when converting all files in a single folder. Default: *.doc

    .PARAMETER TargetFormat
    Word Save AS target format. Currently supported: Default, PDF, XPS, HTML

    .PARAMETER DeleteExistingFiles
    Switch to delete an exiting target file
    .
    .PARAMETER ContinueOnError
    If a file conversion fails, carry on to next file

    .PARAMETER InstancePerFile
    Creates a new instance of the word application per file. Can be useful for larger/slower migrations.

    .EXAMPLE
    Convert all .doc files in E:\temp to Default

    .\Convert-WordDocument.ps1 -SourcePath E:\Temp -IncludeFilter *.doc 

    .EXAMPLE
    Convert all .doc files in E:\temp to XPS

    .\Convert-WordDocument.ps1 -SourcePath E:\Temp -IncludeFilter *.doc -TargetFormat XPS

    .EXAMPLE
    Convert a single document to Word default format

    .\Convert-WordDocument.ps1 -SourcePath E:\Temp\MyDocument.doc
  #>

  [CmdletBinding()]
  Param(
    [string]$SourcePath = '',
    [string]$IncludeFilter = '*.doc',
    [ValidateSet('Default','PDF','XPS','HTML')] # Only some of the supported file formats are currently tested
    [string]$TargetFormat = 'Default',
    [switch]$DeleteExistingFiles,
    [switch]$ContinueOnError,
    [switch]$InstancePerFile
  )

  $ERR_OK = 0
  $ERR_COMOBJECT = 1001 
  $ERR_SOURCEPATHMISSING = 1002
  $ERR_WORDSAVEAS = 1003

  # Define Word target document types
  # Source: https://msdn.microsoft.com/en-us/vba/word-vba/articles/wdsaveformat-enumeration-word  

  $wdFormat = @{
    'Document' = 0 # Microsoft Office Word 97 - 2003 binary file format
    'Template' = 1 # Word 97 - 2003 template format
    'Text' = 2 # Microsoft Windows text format
    'TextLineBreaks' = 3 # 
    'DOSText' = 4 # Microsoft DOS text format
    'DOSTextLineBreaks' = 5 # Microsoft DOS text with line breaks preserved
    'RTF' = 6 # Rich text format (RTF)
    'EncodedText' = 7 # Encoded text format
    'HTML' = 8 # Standard HTML format
    'WebArchive' = 9 # Web archive format
    'FilteredHtml' = 10 # Filtered HTML format
    'XML' = 11 # Extensible Markup Language (XML) format
    'XMLDocument' = 12 # XML document format
    'XMLDocumentMacroEnabled' = 13 # XML document format with macros enabled
    'XMLTemplate' = 14 # XML template format
    'XMLTemplateMacroEnabled' = 15 # XML template format with macros enabled
    'Default' = 16 # Word default document file format. For Word, this is the DOCX format
    'PDF' = 17 # PDF format
    'XPS' = 18 # XML template format
    'FlatXML' = 19 # Open XML file format saved as a single XML file
    'FlatXMLMacroEnabled' = 20 # Open XML file format with macros enabled saved as a single XML file
    'FlatXMLTemplate' = 21 # Open XML template format saved as a XML single file
    'FlatXMLTemplateMacroEnabled' = 22 # Open XML template format with macros enabled saved as a single XML file
    'OpenDocument' = 23 # OpenDocument Text format
    'StrictOpenXMLFormat' = 24 # Strict Open XML document format
  }

  $FileExtension = @{
    'Document' = '.doc'
    'Template' = '.dot'
    'RTF' = '.rtf'
    'HTML' = '.html'
    'Default' = '.docx'
    'PDF' = '.pdf'
    'XPS' = '.xps'
  }

  try{
    # New Word instance
    $WordApplication = New-Object -ComObject Word.Application
  }
  catch {
    Write-Error -Message 'Word COM object could not be loaded'
    Exit $ERR_COMOBJECT
  }


  function Invoke-Word {
  [CmdletBinding()]
  Param(
    [string]$FileSourcePath = '',
    [string]$SourceFileExtension = '',
    [string]$TargetFileExtension = '',
    [int]$WdSaveFormat = 16, # Default to docx
    [switch]$DeleteFile
  )
    if($FileSourcePath -ne '') {
      Write-Output ('Working on {0}' -f $FileSourcePath)
      try {
        $SourceFile = ls $FileSourcePath
        $lastAccessTime = $SourceFile.LastAccessTime
        
        $WordDocument = $WordApplication.Documents.Open($FileSourcePath)
        $NewFilePath = ($FileSourcePath).Replace($SourceFileExtension, $TargetFileExtension)
        $WordDocument.SaveAs([ref] $NewFilePath, [ref]$WdSaveFormat)
        $WordDocument.Close()
        [GC]::Collect()

        $NewFile = ls $NewFilePath
        $NewFile.CreationTime = $SourceFile.CreationTime
        $NewFile.LastWriteTime = $SourceFile.LastWriteTime
        $NewFile.LastAccessTime = $lastAccessTime

        if((Test-Path -Path $NewFilePath) -and $DeleteFile) {          
          $null = Remove-Item -Path $FileSourcePath -Force -Confirm:$false
        }
      }
      catch {
        Write-Error -Message "Error saving document $($FileSourcePath): ´nException: $($_.Exception.Message)"
        $WordDocument.Close()
        [GC]::Collect()
        if($ContinueOnError -ne $true) {
            Exit $ERR_WORDSAVEAS
        }
      }
    }
  }

  if($SourcePath -ne '') {
    $IsFolder = $false
    try {
      $IsFolder = ((Get-Item -Path $SourcePath ) -is [System.IO.DirectoryInfo])
    }
    catch{}
    if($IsFolder) {
      $SourceFiles = Get-ChildItem -Path $SourcePath -Include $IncludeFilter -Recurse
      Write-Verbose -Message ('{0} files found in {1}' -f ($SourceFiles | Measure-Object).Count, $SourcePath)
      foreach($File in $SourceFiles) {
        if($InstancePerFile) {
            try{
                # New Word instance
                $WordApplication = New-Object -ComObject Word.Application
            }
            catch {
                Write-Error -Message 'Word COM object could not be loaded'
                Exit $ERR_COMOBJECT
            }
        }
        if($DeleteExistingFiles) {
            Invoke-Word -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat) -DeleteFile
        } else {
            Invoke-Word -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat) 
        }
      }
    }
    else{
       $File = Get-Item -Path $SourcePath
       if($DeleteExistingFiles) {
            Invoke-Word -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat) -DeleteFile
       } else {
        Invoke-Word -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat)
       }
    }
  }
  else {
    Write-Warning -Message 'No document source path has been provided'
    exit $ERR_SOURCEPATHMISSING
  }
