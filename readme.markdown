# MS Office File Shrink Tool

I created some software which reduces the size of Microsoft Office 2007+ files (.xlsx, .docx, etc.) called Office Shrink. It’s written in batch and makes use of 7-Zip, but was mainly an excuse to play with NSIS (Nullsoft Scriptable Install System). You could build this as a compiled application but batch is more flexible and easier to change for simple tasks. You can download the source code if you want to modify/build the installer yourself.

The tool decompresses an Office 2007 file and repacks it smaller. It then optionally zips the resulting file. The compression is lossless and the file will work as normal. This can be done because the default Office compression is optimised for speed and not for size.

## Installation:

* Run the installer executable from within the zip file
* Click Allow at the UAC prompt (if enabled)
* Read the introduction
* Click OK
* Select if start menu shortcuts are not required
* Click Next
* Change the installation folder if desired
* Click Install
* Click Close

Two new options are added to the windows explorer right click menu for Office 2007 files: “Shrink” and “Shrink and Zip”.

## How to use:

* Ensure the Office 2007 file to compress is not open
* Right click the Office 2007 file in windows explorer
* Click “Shrink” to shrink the file
    * or click “Shrink and Zip” to shrink the file and also zip the shrunk file

## Creative Commons License

This work is licensed under a Creative Commons Attribution 3.0 Unported License.