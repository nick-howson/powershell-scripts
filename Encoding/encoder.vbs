Option Explicit

' Standard housekeeping
Dim objFolder, objFSO, objFile, objShell
Dim colFiles
Dim strSourcePath, strDestinationPath, strItems, strHandbrakePath, strHanbrakeCommand
Dim strMKVExtractPath, strMKVExtractCommand
Dim strMKVMergePath, strMKVMergeCommand
Dim aryFiles
Dim objTextFile
Dim strText

strHandbrakePath = "D:\Applications\Handbrake\HandBrakeCLI.exe"
strMKVExtractPath = "D:\Applications\MKVtoolnix\mkvextract.exe"
strMKVMergePath = "D:\Applications\MKVtoolnix\mkvmerge.exe"

'-o "E:\Unsorted\MakeMKV\test\testing(bwoop).mkv" --chapters "S:\title02.ogm" "E:\Unsorted\MakeMKV\test\testing.mkv"
' Launch standard windows "File Selection" dialouge window for source folder
strSourcePath = SelectFolder( "", "Select source folder" )

' Check that a source folder was selected if not quit
If strSourcePath = vbNull Then
    WScript.Echo "Cancelled"
	WScript.quit
End If

' Launch standard windows "File Selection" dialouge window for destination folder
strDestinationPath = SelectFolder( "", "Select destination folder" )

' Check that a destination folder & source folder were selected if not quit
If ((strSourcePath = vbNull) Or (strDestinationPath = vbNull))  Then
    WScript.Echo "Cancelled"
	WScript.quit
Else
	Set objFSO = CreateObject( "Scripting.FileSystemObject" )
	Set aryFiles = CreateObject( "System.Collections.ArrayList" )
	Set objShell = WScript.CreateObject( "WScript.Shell" )
	Set objFolder = objFSO.GetFolder(strSourcePath)
	Set colFiles = objFolder.Files
	
	' make temp folder
	objFSO.CreateFolder "D:\HandbrakeTemp"
	
	For Each objFile in colFiles
		If UCase(objFSO.GetExtensionName(objFile.name)) = "MKV" Then
		   ' Wscript.Echo objFile.Name
			strItems = strItems & objFile.Name & vbCRLF
			
			' build commandline for mkvextract
			strMKVExtractCommand = """" & strMKVExtractPath & """ chapters -s """ & strSourcePath & "\" & objFile.Name & """ > ""D:\HandbrakeTemp\chapters.ogm"""
			WriteFileText "D:\HandbrakeTemp\bwoop.bat", strMKVExtractCommand 
			
			' build commandline for handbrake
			strHanbrakeCommand = """" & strHandbrakePath & """ -i """ & strSourcePath & "\" & objFile.Name & """ -o ""D:\HandbrakeTemp\temp.mkv"" -f mkv --deinterlace=""slower"" --denoise=""weak"" --strict-anamorphic  -e x264 -q 18 --vfr  -a 1 -E copy  -x ref=6:weightp=1:subq=10:rc-lookahead=10:trellis=2:bframes=5:b-adapt=2:direct=auto:me=tesa:no-dct-decimate=1:b-pyramid=strict --verbose=1 >> D:\HandbrakeTemp\log.txt"
			WriteFileText "D:\HandbrakeTemp\bwoop.bat", strHanbrakeCommand
			
			' build commandline for mkvmerge
			strMKVMergeCommand = """" & strMKVMergePath & """ -o """ & strDestinationPath & "\" & objFile.Name & """ --chapters ""D:\HandbrakeTemp\chapters.ogm"" ""D:\HandbrakeTemp\temp.mkv >> D:\HandbrakeTemp\log.txt"""
			WriteFileText "D:\HandbrakeTemp\bwoop.bat", strMKVMergeCommand
		End If
	Next
	
End If
'objShell.Run "D:\HandbrakeTemp\bwoop.bat", , TRUE
'objFSO.DeleteFolder("D:\HandbrakeTemp")

'"D:\Applications\Handbrake\HandBrakeCLI.exe" -i "E:\Unsorted\MakeMKV\MADMEN_S2_D1\title02.mkv" -c -o "S:\test.mkv"  -f mkv --strict-anamorphic  -e x264 -q 18 --vfr  -a 1 -E copy  -x ref=6:weightp=1:subq=10:rc-lookahead=10:trellis=2:bframes=5:b-adapt=2:direct=auto:me=tesa:no-dct-decimate=1:b-pyramid=strict --verbose=1


Function SelectFolder(StartFolder, strFolderSelectorDescription)
' This function opens a "Select Folder" dialog and will
' return the fully qualified path of the selected folder
'
' Argument:
'     StartFolder    [string]      the root folder where you can start browsing;
'                                  if an empty string is used, browsing starts
'                                  on the local computer
'
' Returns:
' A string containing the fully qualified path of the selected folder
'

    ' Standard housekeeping
    Dim objFolder, objItem, objShell
    
    ' Custom error handling
    On Error Resume Next
    SelectFolder = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, strFolderSelectorDescription, 0, StartFolder )

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function

Function WriteFileText(strTextFilePath, strText)
    Const ForReading = 1
    Const ForWriting = 2
    Const ForAppending = 8
	
    Set objTextFile = objFSO.OpenTextFile(strTextFilePath, ForAppending, True)
    
    ' Write a line.
    objTextFile.WriteLine(strText)

    objTextFile.Close
    'objTextFile.Close

End Function