Dim a
Public uName
Public pName
Dim objFSO
'vars here include a, uName, pName

call application_1InitialSetup()

Function application_1InitialSetup()

'application_1 Setup
'Written by: Cooper Ott

a = MsgBox("Ready to install application_1? Just click OK, and we'll guide you throuh the initial setup.",1+32+0+4096,"Setup: application_1")
'variable a will have value of 1 or 2: 1- OK was clicked 2- Cancel was clicked

If a = 1 Then
'Check if user pushed OK

uName = InputBox("Please enter a username below. It's good to choose something that is identifiable and memorable.","Setup: Username","User404")
'variable uName is Username

pName = InputBox("Please enter a password below. You should choose something with capital letters AND numbers.","Setup: Password","Yeti52")
'variable pName is Password

a = MsgBox("Your current credentials are:" + (Chr(13)) + uName + (Chr(13)) + pName + (Chr(13)) + "Make sure you remember them!",0+64+0+4096,"Setup: Credentials")
'Display uNmae and pName

Set objShell = WScript.CreateObject("WScript.Shell")
'Setting up notifications


Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim this
Set this = objFSO.GetFile("Install.vbs")

Dim loc
Set loc = objFSO.GetFolder(this.ParentFolder)

Dim folderParent
Set folderParent = loc.SubFolders

folderName = "unpack"

Dim nf       
Set nf = folderParent.Add(folderName)

Dim objTextFile
Set objTextFile = objFSO.CreateTextFile("unpack/credentials.txt")
'Create objTextFile, then use objFSO to create the text file that objTextFile is

objTextFile.Close
'Close the ".CreateTextFile" action of objTextFile

Const ForAppending = 8
Set objTextfile = objFSO.OpenTextFile("unpack/credentials.txt",ForAppending,True)
'Use objFSO to .OpenTextFile in Appending mode

objTextfile.WriteLine uName
objTextFile.WriteLine pName
objTextFile.Close
'Write the lines of objTextFile, then close the ".OpenTextFile" action of objTextFile

Set objTextFile = objFSO.CreateTextFile("unpack/application_1.html")
'Create objTextFile, then use objFSO to create the text file that objTextFile is

objTextFile.Close
'Close the ".CreateTextFile" action of objTextFile

Set objTextfile = objFSO.OpenTextFile("unpack/application_1.html",ForAppending,True)
'Use objFSO to .OpenTextFile in Appending mode

objTextfile.WriteLine "<!DOCTYPE html><html lang='en'>"
objTextfile.WriteLine "<head>"
objTextfile.WriteLine "<style>"
objTextfile.WriteLine "p {"
objTextfile.WriteLine "color: red;"
objTextfile.WriteLine "}"
objTextfile.WriteLine "</style>"
objTextfile.WriteLine "</head>"
objTextfile.WriteLine "<body>"
objTextfile.WriteLine "<p>Some text, Lorem Ipsum Dolor Sit Amet</p>"
objTextfile.WriteLine "</body>"

objTextFile.Close
'Write the lines of objTextFile, then close the ".OpenTextFile" action of objTextFile

a = MsgBox("Unpacked")

End If

End Function