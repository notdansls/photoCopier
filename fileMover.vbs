'Declare statements
Dim strVolumeName, strDriveLetter ,strSourcePath, strModifiedYear, strModifiedMonth, strModifiedDay
Dim objFolder


Set objFSO = CreateObject("Scripting.FileSystemObject")
set fsoDrives = objFSO.Drives


'Set variable values
strVolumeName = "NIKON D7000"			' Name of the volume where photos are stored
strSourcePath = ":\DCIM\114D7000"		' This is the folder on the SD Card where photos are stored
photoFolder = "C:\Users\[removed]\OneDrive\Pictures\!inProgress\"

processDrives(strVolumeName)

sub processDrives (strVolumeName)
	for each objDrive in fsoDrives
		if objDrive.VolumeName = strVolumeName then
			strDriveLetter = objDrive.DriveLetter
			processPhotos strDriveLetter, strSourcePath
		end if
	next
end sub

sub processPhotos (strDriveLetter, strSourcePath)
	set objFolder = objFSO.GetFolder(strDriveLetter & strSourcePath)
	set colFiles = objFolder.Files
	for each objFile in colFiles
		
		strModifiedYear = year(FormatDateTime(objFile.DateLastModified,2))
		
		if month(FormatDateTime(objFile.DateLastModified,2)) < 10 then
			strModifiedMonth = "0" & month(FormatDateTime(objFile.DateLastModified,2))
		else
			strModifiedMonth = month(FormatDateTime(objFile.DateLastModified,2))
		end if
		
		if day(FormatDateTime(ObjFile.DateLastModified,2)) < 10 then
			strModifiedDay = "0" & day(FormatDateTime(ObjFile.DateLastModified,2))
		else
			strModifiedDay = day(FormatDateTime(ObjFile.DateLastModified,2))
		end if
		
		strISOModifiedDate = strModifiedYear & "-" & strModifiedMonth & "-" & strModifiedDay
		strDestinationFolder = photoFolder & strISOModifiedDate
		
		if not objFSO.FolderExists(strDestinationFolder) then
			processModifiedDate strDestinationFolder
			if not objFSO.FolderExists(strDestinationFolder) then
				processModifiedDate strDestinationFolder
			end if			
		end if
		
		if objFSO.FolderExists(strDestinationFolder) then
			if objFile.Type = "NEF File" then
				objFile.Move(strDestinationFolder & "\RAW\" & objFile.Name)
			else
				objFile.Move(strDestinationFolder & "\" & objFile.Name)
			end if
		end if
	next	
end sub

sub processModifiedDate (strFullPath)
	do while not objFSO.FolderExists (strFullPath)
		Set objFolder = objFSO.CreateFolder(strFullPath)
		set objFolder = objFSO.CreateFolder(strFullPath & "\RAW")
	loop
end sub
