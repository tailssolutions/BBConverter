
# Blackboard Archive to HTML Converter

## Description
This tool is designed to take portions of a Blackboard course archive and extract the contents, converting it to a single-page HTML file for viewing without the requirement of needing a Server to view. Many portions of the export are based on the [IMS Global Packaging](http://www.imsglobal.org/content/packaging/index.html) standards.

## Usage
The below applies to `bbaExtractorEX.ps1` script

Should be run from the folder containing the PS1 file, will extract to a subfolder of the where the ZIPs are located
```
PS> .\bbaExtractorEX.ps1 -bbArchive ..\ArchiveFile_SP2017_TestCourse.zip
```

The recommended file structure for using this script appears below.

```
Work_Folder/
    ArchiveFile_001.zip     <--- archive from Bb
    ArchiveFile_002.zip

    bbaExtractor/           <--- script folder
        bbaExtractorEX.ps1
        bbaExtractorEX.css
    
    bbaExtraction_FS20xx/   <--- result folder
        /files
```

## Limitations
- Currently, the script does not fully handle calculated scores
- Not all archive formats have been tested

## Change log
- 12/15/2017 - Added `-extractDocuments` option if you want both the HTML report and the document files

    Example Usage:
  
  ```.\bbaExtractorEX.ps1 ..\ArchiveFile_FS2013.MATH.1700.01.zip -extractDocuments```
- 12/15/2017 - Switched to a new version `bbaExtractorEX.ps1` - works directly with zip files
- 07/01/2020 - Modified to parse discussion boards, outputting both original prompt and responses in thread-like format.


## Credits
Majority of the code originally written by Erik 'Jorg' Jorgensen (jorgie@missouri.edu) 2017-2018. Expanded portions written by Chris Potter (chris@tailssolutions.com) 2020. Modified by permission. Code released under MIT license.