'==========================================================================
'
' VBScript Source File -- Created with BurnSoft BurnPad
'
' NAME: DeleteFiles.vbs
' VERSION: 1.0.10
'
' AUTHOR:  BurnSoft, www.burnsoft.net
' DATE  : 1/10/2003
'
' COMMENT: This was created to serach for certain file extensions and delete 
' 			them based on how old they are by a given Number of days. 
'			Great for log maintance it will check all the Sub Directories of the selected root folder'
'==========================================================================
Dim strFile			'-Working File 
Dim strDateCreated  '-Used to get the Last Time a File was modified /Created
Dim strDateOld 		'-Used to get the value of todaydate minus tthe DaysOld Constant
Dim CurrentFile 	'-Working File 
Dim MyFolderList    '-Used to Split the folders in an array
Dim x				'-Count Folder List array
Dim strWorkingDir	'-Current Working Directory
Dim strDateCount    '- To Count files that arex days old.  MOstly used for reporting
Dim Days2Old		'-Used to Convert the Constant Daysold into a negative
Dim strFileType		'-Grab the Array from DayOld
Dim strFileArr		'-File Type Array
Dim ArrayFiles		'-Array of Files Deleted
Dim SplitArray
Dim strNetIQ

CONST RootDirectory = "c:\windows\system32\LogFiles"
CONST FileType = "log"
CONST DaysOld = 30
CONST DeleteAllFiles = "n" 'DeleteAllFiles without DateCheck

'-This Function will get the Date that the selected file was created
'- From the Main Sub this is 4th Called
Function ShowFileInfo(strFile)
	Dim fso,f
	Dim strOld

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFile(strFile)
	'-Use this to get the last time a File was Created
	'strDateCreated = FormatDateTime(f.DateCreated, 2)
	'- Use This to get the Last time a file was modified 
	strDateCreated = FormatDateTime(f.DateLastModified)	 
End Function


'-This Function will delete the selected File 
'- From the Main Sub this is 5th Called
Function DeleteOldFile(strFile)
	Dim fso,f
	Dim strDeletedFile

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set strDeletedFile = fso.GetFile(strFile)
	strDeletedFile.Delete
End Function


'-This is a Main Work Horse that will Compare the Dates and Call for the file to be deleted
'- From the Main Sub this is 3rd Called
Function GetCreatedDate(strFileCreated)
	strFile = RootDirectory & "\" & strWorkingDir & "\" & strFileCreated 
	Call ShowFileInfo(strFile)
	
	'-Formats the Date the File was Created to MM/DD/YY
	GetCurrentDate = FormatDateTime(strdatecreated, 2) 
	
	'-Free Up the GetCurrentDate String
	strDateCreated = GetCurrentDate  
	
	' Now we change the GetCurrentDate string to today's date
	GetCurrentDate = Date  
	
	'- Subtracts the current Date minus 30
	Days2Old = 0 - DaysOld
	StrDateOld = dateadd("d", Days2Old, GetCurrentDate)

   CurrentMonth = dateDiff("d", strDateCreated, GetCurrentDate)
	 
	 if CurrentMonth >= DaysOld then
	
		strDateCount = strDateCount + 1
		ArrayFiles = ArrayFiles + "," + strFile
		'msgbox ArrayFiles
		call DeleteOldFile(strFile)
	end if
End Function
	
'-Get a list of all the subdirectories in the StrDir folder and put into
'- an array with commas to seperate the value	
'- From the Main Sub this is 1st Called
Function GetFolderList(RootDirectory)

	Dim fso, f, fl, s, sf
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(RootDirectory)
	set sf = f.SubFolders
	x = 0
		For Each fl in sf
	    x = x + 1
		s = s & fl.name & ","
	next
	y = 0
	 MyFolderList = split(s, ",")
	 x = x - 1
	 Do until y > x   
	 	strWorkingDir = myfolderlist(y)
	 	If DeleteAllFiles = "y" or DeleteAllFiles = "Y" then
	 		Call GetAllFilesToDelete(strWorkingDir)
	 	else
	 		call getfilelist(strWorkingDir)
	 	end if
	 	y = y + 1
	 loop
end function

'-Get all the files in the Directory
'- From the Main Sub this is 2nd Called
Function GetFileList(strWorkingDir)
	Dim fsof, fi, flf, sf, fc
	Dim strFileSplit
	
	Set fsof = CreateObject("Scripting.FileSystemObject")
	Set fi = fsof.GetFolder( RootDirectory & "\" & strWorkingDir)
	Set fc = fi.Files
	for each flf in fc
		sf = sf & flf.name
		strFileSplit = split(sf, ".")
		if strFileSplit(1) = strFileType then
			CurrentFile = sf
			call GetCreatedDate(CurrentFile)
		end if
		sf = ""
	next
end Function	

'-Get all the files in the Directory
'- From the Main Sub this is 2nd Called if All the Files wish to be deleted
Function GetAllFilesToDelete(strWorkingDir)
	Dim fsof, fi, flf, sf, fc
	Dim strFileSplit
	
	Set fsof = CreateObject("Scripting.FileSystemObject")
	Set fi = fsof.GetFolder( RootDirectory & "\" & strWorkingDir)
	Set fc = fi.Files
	for each flf in fc
		sf = sf & flf.name
		strFileSplit = split(sf, ".")
		if strFileSplit(1) = strFileType then
			CurrentFile = sf
			strFile = RootDirectory & "\" & strWorkingDir & "\" & CurrentFile 
			arrayFiles = ArrayFiles & "," & strFile
			'msgbox Arrayfiles
			call DeleteOldFile(strFile)
			strDateCount = strDateCount + 1
		end if
		sf = ""
	next
end Function


'Sub Main()

'-Get the Number of Extensions used
strFileArr = split(FileType, ",")
strFileType = UBound(strFileArr)

strDateCount = 0
'Set ArrayFiles = nothing
'-Run Option based on NUmber of Extensions
If StrFileType = 0 then
	strFileType = FileType
	If DeleteAllFiles = "y" or DeleteAllFiles = "Y" then
		call GetFolderList(RootDirectory)		
	Else
		call GetFolderList(RootDirectory)		
	End if
else
	For i = 0 to UBound(strFileArr)
		strFileType = strFileArr(i)		
		If DeleteAllFiles = "y" or DeleteAllFiles = "Y" then
			call GetFolderList(RootDirectory)	
		Else
			call GetFolderList(RootDirectory)	
		End if
	Next	
end if	
		
if strDateCount = 0 then
	wscript.echo "0 Files were deleted!"
else
	splitarray = replace(ArrayFiles, ",", Chr(10) & Chr(13))
	strNetIQ = strdatecount & " Files were deleted!" & Chr(10) & Chr(13) & "The Following Files Where Deleted:" & Chr(10) & Chr(13) & splitarray  
	wscript.echo strNetIQ
end if		



'End Sub	
	
	
	
	
	
