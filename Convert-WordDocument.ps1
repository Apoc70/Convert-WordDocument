  <#
    .Synopsis
    PowerShell script to convert Word documents

    .Description

    This script converts Word compatible documents to a selected format utilizing the Word SaveAs function. Each file is converted by a single dedicated Word COM instance.

    The script converts either all documents ina singlefolder of a matching an include filter or a single file.

    Currently supported target document types:
    - Default --> Word 2016
    - PDF
    - XPS
    - HTML

    Author: Thomas Stensitzki
	Minor edits: Gabriel Barcelo
    
    Version 1.2 2022-10-03
    
    .NOTES 
  
    Requirements 
    - Word 2016+/Word 365 installed locally

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0      Initial release
    1.1      Updated Word cleanup code
	1.2      Disabled Macro execution, i.e. AutoOpen()
	         Added -ResetTemplate option

    .LINK
    http://scripts.granikos.eu

    .PARAMETER SourcePath
    Source path to a folder containing the documents to convert or full path to a single document

    .PARAMETER IncludeFilter
    File extension filter when converting all files  in a single folder. Default: *.doc

    .PARAMETER TargetFormat
    Word Save AS target format. Currently supported: Default, PDF, XPS, HTML

    .PARAMETER DeleteExistingFiles
    Switch to delete an exiting target file
	
	.PARAMETER ResetTemplate
	Reset doc to Normal Template. Use in case of very old documents, which may have broken link to template

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
    [switch]$ResetTemplate, # Resets linked template to 'Normal' template
    [switch]$DeleteExistingFiles
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

  function Invoke-Word {
  [CmdletBinding()]
  Param(
    [string]$FileSourcePath = '',
    [string]$SourceFileExtension = '',
    [string]$TargetFileExtension = '',
    [int]$WdSaveFormat = 16, # Default docx
	[switch]$ResetTemplate, # Resets linked template to 'Normal' template
    [switch]$DeleteFile
  )

    if($FileSourcePath -ne '') {

      Write-Output ('Working on {0}' -f $FileSourcePath)
     
      # define variable for Word com object
      $WordApplication = $null

      # try to create a new instance of the COM object
      try{
        # New Word instance
		Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Word\Options" -Name "UseTimeoutForAttachedTemplateLoad" -Value 00000250
        $WordApplication = New-Object -ComObject Word.Application
		#$AutoSecurityDisable = New-Object -ComObject Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable
		# No executa macros com AutoOpen, etc.
		$WordApplication.WordBasic.DisableAutoMacros()
      }
      catch {
        Write-Error -Message 'Word COM object could not be loaded'
		Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Word\Options" -Name "UseTimeoutForAttachedTemplateLoad"
        Exit $ERR_COMOBJECT
      }

      # try to ope the document and save in new format
      try {

        $WordDocument = $WordApplication.Documents.Open($FileSourcePath)

        # Replace the source file extenson with the appropriate target file extension 
        $NewFilePath = ""+($FileSourcePath).Replace($SourceFileExtension, $(" [MIGRAT]"+$TargetFileExtension))
        if($ResetTemplate){
			$WordDocument.AttachedTemplate = "Normal.dotm"
		}
        if((Test-Path -Path $NewFilePath) -and $DeleteFile) {
          # Delete existing file
          $null = Remove-Item -Path $NewFilePath -Force -Confirm:$false
        }

        # Now let's save the document
        $WordDocument.SaveAs([ref] $NewFilePath, [ref]$WdSaveFormat)

      }
      catch {

        # Ooops
        Write-Error -Message "Error saving document$($FileSourcePath): ´nException: $($_.Exception.Message)"
		Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Word\Options" -Name "UseTimeoutForAttachedTemplateLoad"
        Exit $ERR_WORDSAVEAS

      }
      finally{

        # Do some clean up
        # 2019-11-26 Updated
        $WordDocument.Close()
        $WordApplication.Quit()
        [Runtime.Interopservices.Marshal]::ReleaseComObject($WordApplication) | Out-Null

        if(Test-Path variable:global:WordApplication) {
          Remove-Variable -Name WordApplication -Scope Global 4>$Null
        }
        
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
      }
    }
  }

  if($SourcePath -ne '') {

    # Check whether SourcePath is a single file or directory
    $IsFolder = $false
    try {
      $IsFolder = ((Get-Item -Path $SourcePath ) -is [System.IO.DirectoryInfo])
    }
    catch{}

    if($IsFolder) {

      # We need to iterate a source folder
      $SourceFiles = Get-ChildItem -Path $SourcePath -Include $IncludeFilter -Recurse

      Write-Verbose -Message ('{0} files found in {1}' -f ($SourceFiles | Measure-Object).Count, $SourcePath)

      # Let's work on all files
      foreach($File in $SourceFiles) {
		if($ResetTemplate){
          Invoke-Word -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat) -ResetTemplate
		}
		else {
	      Invoke-Word -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat)
		}
      }
    }
    else{
      # It's just a single file

      $File = Get-Item -Path $SourcePath
	  if($ResetTemplate){
        Invoke-Word -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat) -ResetTemplate
	  }
       else {
         Invoke-Word -FileSourcePath $File.FullName -SourceFileExtension $File.Extension -TargetFileExtension $FileExtension.Item($TargetFormat) -WdSaveFormat $wdFormat.Item($TargetFormat)
	   }
    }
  }
  else {
    Write-Warning -Message 'No document source path has been provided'
	Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Word\Options" -Name "UseTimeoutForAttachedTemplateLoad"
    exit $ERR_SOURCEPATHMISSING
  }
