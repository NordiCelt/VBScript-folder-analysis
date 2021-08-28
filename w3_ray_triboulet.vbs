'*************************************************************************
'Script Name: w3_ray_triboulet.vbs
'Author: Ray Triboulet
'Created: 16 August 2021
'Description: ENTD261 Assignment for Week 3: WSH and VBScript
'             - list all file names, sizes, and date created in a given folder
'             - parameter = root folder name (script should validate folder name)
'             - include "flowerbox"
'             - sample run = C:\entd261> .\w3_ray_triboulet.vbs
'                code created with Visual Studio Code CE and run within powershell terminal
'             - sample run = C:\ENTD261>w3_ray_triboulet.vbs > results.txt
'                alternatively run on CMD Prompt with optional results.txt file
'*************************************************************************

Dim fso, f, colFiles, NewFile                            'creates variables

Set fso = CreateObject("Scripting.FileSystemObject")     'sets fso var to creating a file system object
Set f = fso.GetFolder("C:\ENTD261")                      'sets f var to collect data from folder
Set colFiles = f.Files                                   'creates temp collected files data for the script to work with
Set NewFile = fso.CreateTextFile("results.txt")          'creates a text file to write into

For Each f in colFiles                                   'for loop to collect and write desired data to file
    NewFile.WriteLine(f.Name)                             'write file name to "results.txt"
    NewFile.WriteLine(f.Size)                             'write file size to "results.txt"
    NewFile.WriteLine(f.DateCreated)                      'write file date to "results.txt"
    NewFile.WriteLine(f.Type)                             'write file type to "results.txt"
Next
