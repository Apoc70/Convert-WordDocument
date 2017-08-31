# Convert-WordDocument.ps1
PowerShell script to convert Word documents

## Description
This script converts Word compatible documents to a selected format utilizing the Word SaveAs function. Each file is converted by a single dedicated Word COM instance.

The script converts either all documents ina singlefolder of a matching an include filter or a single file.

Currently supported target document types:
- Default --> Word 2016
- PDF
- XPS
- HTML

## Parameters

### Parameter SourcePath
Source path to a folder containing the documents to convert or full path to a single document

### Parameter IncludeFilter
File extension filter when converting all files  in a single folder. Default: *.doc

### Parameter TargetFormat
Word Save AS target format. Currently supported: Default, PDF, XPS, HTML

### Parameter DeleteExistingFiles
Switch to delete an exiting target file

## Examples
```
.\Convert-WordDocument.ps1 -SourcePath E:\Temp -IncludeFilter *.doc 
```
Convert all .doc files in E:\temp to Default

```
.\Convert-WordDocument.ps1 -SourcePath E:\Temp -IncludeFilter *.doc -TargetFormat XPS
```
Convert all .doc files in E:\temp to XPS

```
.\Convert-WordDocument.ps1 -SourcePath E:\Temp\MyDocument.doc
```
Convert a single document to Word default format

## Note
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery
Download and vote at TechNet Gallery
* 

## Credits
Written by: Thomas Stensitzki

## Social 

* My Blog: http://justcantgetenough.granikos.eu
* Twitter: https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github: https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog: http://blog.granikos.eu/
* Website: https://www.granikos.eu/en/
* Twitter: https://twitter.com/granikos_de