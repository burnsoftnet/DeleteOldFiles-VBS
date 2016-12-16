'==========================================================================
'
' VBScript Source File -- Created with BurnSoft BurnPad
'
' NAME: DeleteFiles.vbs
' VERSION: 2.0.1
'
' AUTHOR:  BurnSoft, www.burnsoft.net
' DATE  : 9/10/2004
'
' COMMENT: This script is the next best version of the DeleteFiles.vbs script
'		Unlike the deletefiles job, this one will go through all levels for 
'		the folder structure looking for old files.
'
'==========================================================================

Const strPath="C:\temp\"     ' Target Directory
Const strPattern="*.log"     ' Extension Pattern
Const DeleteFiles="y"     'Delete Files (y/n)
Const DaysOld=10     ' Delete Files older then x many days
Const UseMessages="y"     ' Send Messages to console
Const AlertIfNothing="n"     ' Send Message is nothing is listed (y/n)
Const UseMyExemptFile="n"     ' Use File Exemption (y/n), Skip folder if file found
Const MyExempt="Health.txt"     ' Exempt file to look for
Const USEMYEXEMPTFOLDER="n"     ' Use folder Exemption?
Const EXEMPTFOLDER="C:\Temp\VB"     ' Exempt Folders List (comma seperated)
Dim FolderCount
Dim sMessage
Dim NL
Const MessageMax=6300
'=======================================================================
Function isFolderExempt(strFolder)
Dim strSplit
Dim intBound
Dim bAns
bAns = False
If USEMYEXEMPTFOLDER = "y" Then
	strSplit = split(EXEMPTFOLDER,",")
	intbound = ubound(strsplit)
	If intbound > 0 Then
		For i = 0 To intbound
			If trim(ucase(strFolder)) = trim(ucase(strsplit(i))) Then
				bAns = True
				Exit For
			Else
				bAns = False
			End If
		Next
	Else
		If trim(ucase(strFile)) = trim(ucase(MyExempt)) Then
			bAns = True
		Else
			bAns = False
		End If	
	End If
End If
isFolderExempt = bans
End Function
'=======================================================================
Function IsExempt(strFile)
Dim strSplit
Dim intBound

If len(MyExempt) > 0 And UseMyExemptFile="y" Then
	strSplit = split(MyExempt,",")
	intbound = ubound(strsplit)
	
	If intbound > 0 Then
		For i = 0 To intbound
			If trim(ucase(strFile)) = trim(ucase(strsplit(i))) Then
				IsExempt = True
				Exit Function
			Else
				IsExempt = False
			End If
		Next
	Else
		If trim(ucase(strFile)) = trim(ucase(MyExempt)) Then
			IsExempt = True
			Exit Function
		Else
			IsExempt = False
		End If	
	End If
Else
	IsExempt = False
End If
End Function
'=======================================================================
Function DeleteOldFile(strFile)
	Dim fso,f
	Dim strDeletedFile
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set strDeletedFile = fso.GetFile(strFile)
	strDeletedFile.Delete
End Function
'=======================================================================
Function ShowFileInfo(strFile)
	Dim fso,f
	Dim strOld
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFile(strFile)
	ShowFileInfo = FormatDateTime(f.DateLastModified)	 
End Function
'=======================================================================

Function GetCreatedDate(strFile)
	Dim strDateCreated, GetCurrentDate
	strDateCreated = FormatDateTime( ShowFileInfo(strFile), 2)
	GetCurrentDate = Date  
	Days2Old = 0 - DaysOld
	StrDateOld = dateadd("d", Days2Old, GetCurrentDate)
    CurrentMonth = dateDiff("d", strDateCreated, GetCurrentDate)
	 
	 If CurrentMonth >= DaysOld Then
 		If len(smessage) >= MessageMax Then
 			If usemessages = "y" Then
				wscript.echo "Deleted Files Report:" & NL & smessage
			End If
			smessage = ""
 		End If
	 	If smessage = "" then
			smessage = "The Following Files where greater then " & DaysOld & " days old and  where deleted." & Chr(10) & Chr(13)
		End If
		strDateCount = strDateCount + 1
		ArrayFiles = ArrayFiles + "," + strFile
	 	If DeleteFiles="y" Then Call DeleteOldFile(strFile)
		smessage = smessage & strfile & Chr(10) & Chr(13)
	End If
End Function
'=======================================================================
Function CountStuff(path,subfolders, Size, ReportType)
Dim fso
Dim fs
Dim f
Dim cSubFolders
Dim CountArchive
Dim totalBytes
err.clear
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set fs = fso.GetFolder(path)
If isFolderExempt(path) Then Exit Function
cSubFolders = cSubFolders + fs.subfolders.count

For Each file In fs.Files
	Dim ext1
	Dim ext2
	If IsExempt(file.Name) = False Then
		ext1 = Len(file.Name) - InStrRev(file.Name, ".")
		ext2 = Len(strpattern) - InStrRev(strpattern, ".")
		If Right(UCase(file.Name), ext1) = Right(UCase(strpattern), ext2) Then
			CountArchive = CountArchive + 1
			totalBytes = totalBytes + file.Size
			GetCreatedDate(path & "\" & file.Name)
		ElseIf strpattern = "*.*" Then
			CountArchive = CountArchive + 1
			totalBytes = totalBytes + file.Size
			GetCreatedDate(path & "\" & file.Name)
		End If
	End If
Next

Dim Folder

For Each folder In fs.subfolders
	path = folder.path
	CountArchive = CountArchive + CountStuff(path,subfolders,size,"FILES")
	cSubFolders = cSubFolders + fs.subfolders.count
Next
Select Case ReportType
	Case "FILES"
		CountStuff = CountArchive
	Case "FOLDERS"
		CountStuff = cSubFolders
	Case "BYTES"
		CountStuff = totalBytes
End select
End Function
'=======================================================================
sMessage = ""
FolderCount = ""
NL = Chr(10) & Chr(13)
If CountStuff(strpath,0,0, "FILES") > 0 Then
	If smessage <> "" Then
		If usemessages = "y" Then
			WScript.Echo "Deleted Files Report" & NL & smessage
		End If
	Else
		If AlertIfNothing = "y" Then
			WScript.Echo "No Files Deleted Report" & NL & "No Files Deleted Report"
		End If
	End if
End If
sMessage = ""
Set sMessage = Nothing
FolderCount = ""