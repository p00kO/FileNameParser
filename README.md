# FileNameParser
FileNameParser is an application built to extend image capture and image processing software. In it's current version, it allows a user to add predefined metadata to a .tif file by parsing the file name for strings in a lookup table. The current version wraps a microscope application (AmScope), watching for the user to save images to disk. When that happens, it reopens the images and adds metadata about the microscope, the camera, the objective and relay lenses and a calculated pixel pitch to the .tif file captured by the imaging camera software. 

## List of Files:
- **RegMod:**
Add this key to the registry to launch the FileNameParser every time the microscope application is launched. The test1.bat file will be launched as a debugger instead of the microscope application (amscope.exe).
- **test1.bat:** 
Called when a user starts the microscope application, it attempts to launch the powershell parent process script MicroscopeApp2.ps in administrator mode.
- **MicroscopeApp2.ps:** 
Launched in administrator mode, this script first removes the ...\Image File Execution Options\... registry key used to redirect the call to our test1.bat file and starts the microscope app. It also starts the FileNameParser C# application. MicoscopeApp2.ps does not complete until the amscope.exe application terminates. To change the application that will be watched, modify the $micAppName variable and the micAppPath variable. On install, ensure that the localWatcherDir path matches the path where the FileNameParser project was saved.
- **WindowsFormApp1.sln:**
This is the C# application which watches the application for .tif file write to disc. It uses the Windows ETW framework to monitor the process ID of the microscope application passed to it when it started from the MicroscopeApp2.ps1 script. Every time a user saves a .tif file with the amscope.exe application, a callback function is called about 1 second after the user saved the file, providing the file path.  The filename is parsed from the provided file path for the objective lens magnification and relay lens. The filename should end with and _---x---r where ---x is the objective lens magnification and ---r is the relay lens as listed in the **turretwatcher_*%datestamp%*.xml** file. For example; ( ex: filename_5x1r.tif ). The **turretwatcher.config** file stores the name of the current .xml calibration file. 
## Notes on Permissions:
Because the application will modify the registry and uses the ETW framework, it will only work with administrator privileged. There may be issues with some anti-virus software scans.
