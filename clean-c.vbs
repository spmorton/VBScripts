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
' --------------------------------------------------------' 

Wscript.Echo "Running script clean-c.vbs"

' Setup the shell object to execute external commands
Set WshShell = WScript.CreateObject("WScript.Shell")

On Error Resume Next 

' Delete specific files
call delete_file("C:\Windows\MEMORY.DMP")

' Delete directory trees
call delete_dir_tree("C:\temp\")
call delete_dir_tree("C:\_temp\")
call delete_dir_tree("C:\Windows\Logs\WindowsServerBackup")
call delete_dir_tree("C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\Quarantine\")
call delete_dir_tree("C:\Documents and Settings\All Users\Application Data\Symantec\LiveUpdate\Downloads\")
call delete_dir_tree("C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\Logs\")
call delete_dir_tree("C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\Quarantine\")
call delete_dir_tree("C:\ProgramData\Symantec\LiveUpdate\Downloads\")
call delete_dir_tree("C:\ProgramData\Symantec\Symantec Endpoint Protection\Logs\")
call delete_dir_tree("C:\ProgramData\Symantec\Symantec Endpoint Protection\Quarantine\")
call delete_dir_tree("C:\Windows\System32\CCM\Cache\")
call delete_dir_tree("C:\Windows\System32\CCM\Logs\")
call delete_dir_tree("C:\Windows\SysWOW64\CCM\Cache\")
call delete_dir_tree("C:\Windows\SysWOW64\CCM\Logs\")
call delete_dir_tree("C:\ProgramData\Microsoft\WIndows\WER\ReportQueue")
call delete_dir_tree("C:\ProgramData\Microsoft\WIndows\WER\ReportQueue")
call delete_dir_tree("C:\Windows\SoftwareDistribution\Download")
call delete_dir_tree("C:\Windows\Minidump")
call delete_dir_tree("C:\WINDOWS\system32\LogFiles\W3SVC1")

' Delete directory tree by date
call delete_dir_tree_by_date("C:\Program Files\Common Files\Symantec Shared\VirusDefs", 30)
call delete_dir_tree_by_date("C:\inetpub\logs\LogFiles\W3SVC1", 14)
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

' Clear the recylcle.bin on 2008
call delete_file_tree ("$recycle.bin")

' Clear the recylcle.bin on 2000/2003
call delete_file_tree ("recycler")


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

' Stop specified service
function stop_service (servicename)
	call WshShell.Run("net stop " & servicename, 1, true)
end function

' Start specified service
function start_service (servicename)
	call WshShell.Run("net start " & servicename, 1, true)
end function

' Delete a specific file
function delete_file (filename)
	Set fSystem = CreateObject("Scripting.FileSystemObject")
	If (fSystem.FileExists(filename)) Then
		call fSystem.DeleteFile(filename, true)
	End If
end function
	
' Delete everything under a named path
function delete_at_path (path)

	On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) then
		Wscript.Echo "Deleting files at " & path	
		Set directory = fSystem.GetFolder(path) 
		Set fileSet = directory.Files 

		For each child in fileSet
			child.Delete(True) 
		Next 
	end if
end function

' Delete everything under a named path that is older than the reference days given
function delete_by_date( path, date_ref )

	On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject") 
	if (fSystem.FolderExists(path) then
		Set directory = fSystem.GetFolder(path) 
		Set fileSet = directory.Files 

		Wscript.Echo "Deleting files at " & path	

		For each child in fileSet
			If child.DateLastModified < (Date() - date_ref) Then 
				call fSystem.DeleteFile(child, true) 
			End If 
		Next 
	end if
end function

' Delete every file in a given path including files in sub-folders but not the directory tree
function delete_file_tree( path )

	On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) then
		Set directory = fSystem.GetFolder(path)

		call delete_at_path(path)
		
		For Each subFolder in directory.SubFolders
			call delete_at_path(subFolder.path)
			call delete_file_tree(subFolder.path)
		Next
	end if
end function

' Delete everything under a given path including sub-folders
function delete_dir_tree( path )

	On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) then
		Set directory = fSystem.GetFolder(path)

		call delete_at_path(path)
		
		For Each subFolder in directory.SubFolders
			call delete_at_path(subFolder.path)
			call delete_dir_tree(subFolder.path)
			fSystem.DeleteFolder(subFolder.path)
		Next
	end if
end function

' Delete everything under a given path including sub-folders older than given days
function delete_dir_tree_by_date( path, date_ref )

	On Error Resume Next 

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) then
		Set directory = fSystem.GetFolder(path)

		call delete_by_date(path, date_ref)

		For Each subFolder in directory.SubFolders
			call delete_by_date(subFolder.path, date_ref)
			call delete_dir_tree_by_date(subFolder.path, date_ref)
			fSystem.DeleteFolder(subFolder.path)
		Next
	end if
end function

' Delete files based on given ext name Ex ".bin"
function delete_files_by_ext(path, ext)

	Set fSystem = CreateObject("Scripting.FileSystemObject")
	if (fSystem.FolderExists(path) then
		Set Folder = fSystem.GetFolder(path)

		For Each File In Folder.Files
			If Right(File.name, 4) = ext Then
				fSystem.DeleteFile(fldPath & "\" & File.name)
			End If
		Next
	end if
end function
