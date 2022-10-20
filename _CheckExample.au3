
;~ ShellExecuteWait(@ScriptDir & '\7za.exe', 'a "' & 'C:\Archive.7z" -p"Мой пароль" -mhe -mx9 "' & 'C:\P1005.log"', '', '', @SW_HIDE)

Select
	Case FileExists('C:\Program Files\7-Zip\7z.exe')
		ShellExecuteWait('C:\Program Files\7-Zip\7z.exe', 'a "' & 'C:\Archive.7z" -p"Мой пароль" -mhe -mx9 "' & 'C:\P1005.log"', '', '', @SW_HIDE)
	Case FileExists('C:\Program Files (x86)\7-Zip\7z.exe')
		ShellExecuteWait('C:\Program Files (x86)\7-Zip\7z.exe', 'a "' & 'C:\Archive.7z" -p"Мой пароль" -mhe -mx9 "' & 'C:\P1005.log"', '', '', @SW_HIDE)
	Case Else
		MsgBox(4096, '', @ProgramFilesDir & '1')
EndSelect