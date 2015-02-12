' clean-c.vbs
' VBScript to clean the c-drive of common junk that causes space issues on servers
' Author Scott Morton
' September 20, 2013
' // --Version 0.9
' October 29, 2013
' Version 1.0
' Script has been used in production and new functionality added as of this writing.
' Added logic to verify paths exist prior to performing operations
' Added specifics for other logs and recycle bins
'
' 10/21/2014
' Version 1.1
' Added script logging
' Added logic to terminate function if action taken on anything not on C:
' Added logic to terminate function if specified path is an alias to another location(link,hlink,junction)
' Corrected logic for recycle bins

' 11/21/2014
' Version 1.2
' Corrected DFBE to actually perform the requested action
' Expanded capabilities of DFBE to take a long extension name
' Added logging for the starting of each function
'
' 12/23/2014
' Version 1.2.1
' added additional directory to check for dated log files
' call delete_dir_tree_by_date("C:\WINDOWS\system32\LogFiles", 30)
'
' 1/23/2015
' Version 1.2.2
' Added logging of user executing the script
' converted script version to single variable for ease of maintenance
' Added feeble attempt to detect log tampering
' --------------------------------------------------------' 

Script_Version = "1.2.2"

'Setup the log file
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
if (Not ObjFSO.FolderExists("c:\clean-c.log")) then
	ObjFSO.CreateFolder("c:\clean-c.log")
End If

TimeStamp = Now 
TimeStamp = Replace(TimeStamp, "/" ,"-") 
TimeStamp = Replace(TimeStamp ,":","_") 

Set objLog = objFSO.CreateTextFile("c:\clean-c.log\" & Timestamp & ".txt")

' End log file setup

' !!!!!!!!! Operations begin here !!!!!!!!!!!!
' Setup the shell object to execute external commands
Set WshShell = WScript.CreateObject("WScript.Shell")

Wscript.Echo "Running script clean-c.vbs"
Wscript.echo "Version: " & Script_Version

' Capture user executing the script
UserName = WshShell.ExpandEnvironmentStrings( "%USERNAME%" )

objLog.WriteLine("clean-c.vbs Version: " & Script_Version)
objLog.WriteLine("Time Stamp: " & Now)
objLog.WriteLine("Execution Stamp: " & Timer)
objLog.WriteLine("Executed by: " & Username)

' Maintain the log file directory
call delete_by_date("c:\clean-c.log", 60)


On Error Resume Next 

' Delete specific files
call delete_file("C:\Windows\MEMORY.DMP")

' Clear Recycle Bin for all users
Set fSystem = CreateObject("Scripting.FileSystemObject")
If (fSystem.FolderExists("c:\recycler")) Then
	call delete_file_tree("c:\recycler")
End If

If (fSystem.FolderExists("c:\$Recycle.Bin")) Then
	call delete_file_tree("c:\$Recycle.Bin", true)
End If


' Delete directory trees
' Removed for Cerner specific reasons. Do not re-enable; call delete_dir_tree("C:\temp\")
' Removed for Cerner specific reasons. Do not re-enable; call delete_dir_tree("C:\_temp\")
call delete_dir_tree("C:\Windows\Logs\WindowsServerBackup")
call delete_dir_tree("C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\Quarantine\")
call delete_dir_tree("C:\Documents and Settings\All Users\Application Data\Symantec\LiveUpdate\Downloads\")
call delete_dir_tree("C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\Logs\")
call delete_dir_tree("C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\Quarantine\")
call delete_dir_tree("C:\ProgramData\Symantec\LiveUpdate\Downloads\")
call delete_dir_tree("C:\ProgramData\Symantec\Symantec Endpoint Protection\Logs\")
call delete_dir_tree("C:\ProgramData\Symantec\Symantec Endpoint Protection\Quarantine\")
call delete_dir_tree("C:\Windows\System32\CCM\Cache\")
call delete_dir_tree("C:\Windows\System32\CCM\temp\")
call delete_dir_tree("C:\Windows\System32\CCM\Logs\")
call delete_dir_tree("C:\Windows\SysWOW64\CCM\Cache\")
call delete_dir_tree("C:\Windows\SysWOW64\CCM\Logs\")
call delete_dir_tree("C:\ProgramData\Microsoft\WIndows\WER\ReportQueue\")
call delete_dir_tree("C:\ProgramData\Microsoft\WIndows\WER\ReportQueue\")
call delete_dir_tree("C:\Windows\SoftwareDistribution\Download\")
call delete_dir_tree("C:\Windows\Minidump\")
call delete_dir_tree("C:\WINDOWS\system32\LogFiles\W3SVC1\")
call delete_dir_tree("C:\Program Files\Common Files\microsoft shared\Web Server Extensions\14\LOGS\")
call delete_dir_tree("C:\ProgramData\Microsoft\Windows\WER\ReportQueue\")


' Delete files from directory tree
call delete_file_tree ("C:\WINDOWS\system32\CCM\Inventory\temp\")

' Delete directory tree by date
call delete_dir_tree_by_date("C:\Program Files\Common Files\Symantec Shared\VirusDefs", 30)
call delete_dir_tree_by_date("C:\inetpub\logs\LogFiles", 14)
call delete_dir_tree_by_date("C:\WINDOWS\system32\LogFiles", 30)
call delete_dir_tree_by_date("C:\Radius_Logs", 30)

' clear IBM Director 6.x log files
call stop_service ("wmicimserver")
call stop_service ("cimlistener")
call stop_service ("TWGIPC")
call delete_file("C:\Program Files\IBM\Director\data\esntevt.dat")
call delete_file("C:\Program Files\IBM\Director\log\CimUrlCgi.log")
call delete_file("C:\Program Files\IBM\Director\log\twgescli.bak")
call delete_file("C:\Program Files\IBM\Director\log\twgescli.log")
call start_service ("cimlistener")
call start_service ("wmicimserver")
call start_service ("TWGIPC")


'Purge Cache
call WshShell.Run("sfc /purgecache", 1, true)

'Remove profiles that have not been accessed in 180 days
call WshShell.Run("C:\l2serversupport\delprof1.exe /q /i /d:180", 1, true)

'Clean up WINSXS if possible

call WshShell.Run("DISM.exe /online /Cleanup-Image /spsuperseded", 1, true)
call WshShell.Run("Cmpcln.exe", 1, true)
call WshShell.Run("Vsp1cln.exe", 1, true)

call WshShell.Run("net stop trustedinstaller", 1, true)

call WshShell.Run("takeown /f %windir%\winsxs\ManifestCache\*.bin", 1, true)
call WshShell.Run("cacls %windir%\winsxs\ManifestCache\*.bin /t /e /g %username%:F", 1, true)

call delete_files_by_ext("%windir%\winsxs\ManifestCache", ".bin")

call WshShell.Run("net start trustedinstaller", 1, true)
' end WINSXS


'!!!!!!!!!!!!!!! Functions Start Here !!!!!!!!!!!!!!!!!!!!



' Stop specified service
function stop_service (servicename)
	On Error Resume Next 
	objLog.writeline("Stopping service - " & servicename)
	call WshShell.Run("net stop " & servicename, 1, true)
end function

' Start specified service
function start_service (servicename)
	On Error Resume Next 
	objLog.writeline("Starting service - " & servicename)
	call WshShell.Run("net start " & servicename, 1, true)
end function

' Check to make sure the folder is not an alias to another place
' Returns true if the folder is not an alias
function is_folder(path)

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	Set folder = fSystem.GetFolder(path)
	attr = folder.attributes And 1024
	if ( attr = 0 ) then
		is_folder = true
		Exit function
	end if
	
	objLog.writeline("folder is an aliased folder - " & path)
	is_folder = false

end function

' Check to make sure that the path is not on an aliased path
' Checks each subfolder up to the root of the drive
function check_path(path)

	if (instrrev(path,".") = 0) then
		pos = len(path)
	else	
		pos = instrrev(path,"\")
	end if
	do
		if NOT(is_folder(left(path,pos))) then
			Exit Do
		end if
		pos = instrrev(path,"\",pos - 1)
	loop while( pos > 0 )
	if ( pos = 0 ) then
		check_path = true
		Exit function
	else
		objLog.writeline("file is contained in an aliased folder at " & left(path,pos))
		check_path = false
	end if

end function


' Delete a specific file
function delete_file (filename)
	objLog.WriteLine("starting DF")

	On Error Resume Next 
	Set fSystem = CreateObject("Scripting.FileSystemObject")
	If (fSystem.FileExists(filename) And (lcase(left(filename,3)) = "c:\") And check_path(filename) ) Then
		objLog.WriteLine("DF is Deleting file - " & filename )
		call fSystem.DeleteFile(filename, true)
		delete_file = true
		Exit function
	else
		objLog.WriteLine("DF is aborting delete of - " & filename )
		delete_file = false
		Exit function
		
	End If
end function
	
' Delete everything under a named path
function delete_at_path (path)
	objLog.WriteLine("starting DAP")

	On Error Resume Next 
	
	Set fSystem = CreateObject("Scripting.FileSystemObject")
	
	if (fSystem.FolderExists(path) And (lcase(left(path,3)) = "c:\") And check_path(path)) then
		objLog.WriteLine("DAP is Deleting files at - " & path )
		'Wscript.Echo "Deleting files at " & path	
		Set directory = fSystem.GetFolder(path) 
		Set fileSet = directory.Files 

		For each child in fileSet
			if (check_path(child)) then
				objLog.WriteLine("DAP is Deleting file - " & child )
				child.Delete(True)
			end if
		Next
		delete_at_path = true
		Exit function
	else
		objLog.WriteLine("DAP is aborting deleting files at - " & path )
		delete_at_path = false
	end if
end function

' Delete everything under a named path that is older than the reference days given
function delete_by_date( path, date_ref )
	objLog.WriteLine("starting DBD")

	On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject") 
	if (fSystem.FolderExists(path) And (lcase(left(path,3)) = "c:\") And chsck_path(path)) then
		Set directory = fSystem.GetFolder(path) 
		Set fileSet = directory.Files 

		objLog.WriteLine("DBD is Deleting files at - " & path )
		'Wscript.Echo "Deleting files at " & path	

		For each child in fileSet
			If child.DateLastModified < (Date() - date_ref) Then
				objLog.WriteLine("DBD is Deleting file - " & child )
				call fSystem.DeleteFile(child, true) 
			End If 
		Next
		delete_by_date = true
		Exit function
	else
		objLog.WriteLine("DBD is aborting deleting files at - " & path )
		delete_by_date = false
	end if
end function

' Delete every file in a given path including files in sub-folders but not the directory tree
function delete_file_tree( path )
	objLog.WriteLine("starting DFT")

	On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) And (lcase(left(path,3)) = "c:\") And check_path(path)) then
		Set directory = fSystem.GetFolder(path)

		if( delete_at_path(path) ) then
			For Each subFolder in directory.SubFolders
				if not( delete_at_path(subFolder.path) ) then
					delete_file_tree = false
					Exit function
				end if
					
				if not( delete_file_tree(subFolder.path) ) then
					delete_file_tree = false
					Exit function
				end if
			Next

			delete_file_tree = true
			Exit function

		else
			delete_file_tree = false
		end if
	end if
end function

' Delete everything under a given path including sub-folders
function delete_dir_tree( path )
	objLog.WriteLine("starting DDT")

	'On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) And (lcase(left(path,3)) = "c:\") And check_path(path) ) then
		Set directory = fSystem.GetFolder(path)

		if ( delete_at_path(path) ) then
			For Each subFolder in directory.SubFolders
				if not( delete_at_path(subFolder.path) ) then
					delete_dir_tree = false
					Exit function
				end if
				if not( delete_dir_tree(subFolder.path) ) then
					delete_dir_tree = false
					Exit function
				end if
				objLog.WriteLine("Deleting folder " & subFolder.path )
				call fSystem.DeleteFolder(subFolder.path)
			Next
			delete_dir_tree = true
			Exit function
		else
			delete_dir_tree = false
		end if
	end if
end function

' Delete everything under a given path including sub-folders older than given days
function delete_dir_tree_by_date( path, date_ref )
	objLog.WriteLine("starting DDTBD")

	On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) And (lcase(left(path,3)) = "c:\") And check_path(path) ) then
		Set directory = fSystem.GetFolder(path)

		if ( delete_by_date(path, date_ref) ) then
			For Each subFolder in directory.SubFolders
				if not( delete_by_date(subFolder.path, date_ref) ) then
					delete_dir_tree_by_date = false
					Exit function
				end if
				if not( delete_dir_tree_by_date(subFolder.path, date_ref) ) then
					delete_dir_tree_by_date = false
					Exit function
				end if
				objLog.WriteLine("DDTBD is Deleting folder - " & subFolder.path )
				call fSystem.DeleteFolder(subFolder.path)
			Next
			delete_dir_tree_by_date = true
			Exit function
		else
			objLog.WriteLine("DDTBD is aborting deleting folder - " & subFolder.path )
			delete_dir_tree_by_date = false
			Exit function
		end if
	else
		objLog.WriteLine("DDTBD is aborting deleting folders at - " & subFolder.path )
		delete_dir_tree_by_date = false
	end if
end function

' Delete files based on given ext name Ex ".bin"
function delete_files_by_ext(path, ext)
	objLog.WriteLine("starting DFBE")

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) And (lcase(left(path,3)) = "c:\") And check_path(path) ) then
		Set directory = fSystem.GetFolder(path)
		
			objLog.WriteLine(directory.Name)
		
		Set fileSet = directory.Files
		
		For Each File In fileSet
			pos = instrrev(File,".")
			If (right(File, len(File) - (pos - 1)) = ext) Then
				objLog.WriteLine("DFBE is Deleting file " & File )
				call fSystem.DeleteFile(File, true)
			End If
		Next
	else
		objLog.WriteLine("DFBE is aborting deleting files at - " & path )

	end if
end function
