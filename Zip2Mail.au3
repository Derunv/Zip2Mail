#cs ----------------------------------------------------------------------------
	Standart Copy  Folder

 AutoIt Version: 3.3.12.0
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

#include <GUIConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <FileOperations.au3>

FileInstall( ".\App\Pic\Fon.jpg" , "C:\Temp\Fon.jpg" , 1)
FileInstall( ".\App\Pic\Done.jpg" , "C:\Temp\Done.jpg" , 1)
FileInstall( ".\App\Pic\Wait.jpg" , "C:\Temp\Wait.jpg" , 1)
FileInstall( ".\App\Pic\Span.jpg" , "C:\Temp\span.jpg" , 1)
FileInstall( ".\App\Pic\\HowTo.jpg" , "C:\Temp\\HowTo.jpg" , 1)

Global $file = "C:\Temp\Archive.7_"

$hGUI = GUICreate("Zip2Mail", 385, 263, 391, 450, -1, BitOR($WS_EX_ACCEPTFILES, $WS_EX_TOPMOST))
GUISetBkColor(0xFFFFFF)
$Button1 = GUICtrlCreateButton("About", 10, 190, 70, 25)
GUICtrlSetState (-1, $GUI_DROPACCEPTED)
$Button2 = GUICtrlCreateButton("How to Use", 300, 190, 80, 50)
$Pic1 = GUICtrlCreatePic("C:\Temp\Fon.jpg", 5, 5, 375, 177)
GUICtrlSetState (-1, $GUI_DROPACCEPTED)
GUICtrlSetCursor (-1, 2)
GUISetState(@SW_SHOW)
GUIRegisterMsg ($WM_DROPFILES, "WM_DROPFILES_FUNC")
GUICtrlCreateLabel( "Created by Derun Vitaliy", 270, 250)
GUICtrlSetFont(-1, 7)

;~ $About = GUICreate("About", 450, 434, 192, 124)

Global $gaDropFiles[1], $str = ""

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
        Case $GUI_EVENT_DROPPED
			GUICtrlSetImage($Pic1, "C:\Temp\Wait.jpg")
			WinSetOnTop($hGUI, '',0)
;~ 			GUICtrlSetData($Button2, "Close")
			To_ZIP()
		Case $Button1
			MsgBox(0+0+0+0, 'About Zip2Mail', 'Zip2Mail' & @CRLF & 'Version 1.0.2' & @CRLF & '    March 13 2015 16:06:38 ' & @CRLF & 'Created by Derun Vitaliy' & @CRLF)
		Case $Button2
			ShellExecuteWait("C:\Temp\HowTo.jpg")
	EndSwitch
WEnd

Func WM_DROPFILES_FUNC($hWnd, $msgID, $wParam, $lParam)
    Local $nSize, $pFileName
    Local $nAmt = DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", 0xFFFFFFFF, "ptr", 0, "int", 255)
    For $i = 0 To $nAmt[0] - 1
        $nSize = DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", $i, "ptr", 0, "int", 0)
        $nSize = $nSize[0] + 1
        $pFileName = DllStructCreate("char[" & $nSize & "]")
        DllCall("shell32.dll", "int", "DragQueryFile", "hwnd", $wParam, "int", $i, "ptr", DllStructGetPtr($pFileName), "int", $nSize)
        ReDim $gaDropFiles[$i+1]
        $gaDropFiles[$i] = DllStructGetData($pFileName, 1)
        $pFileName = 0
    Next
EndFunc

Func TO_ZIP()
	FileDelete('C:\Temp\Archive.7_')
    Select
		Case FileExists('C:\Program Files\7-Zip\7z.exe')
			Switch UBound($gaDropFiles)
				Case 1
					$aExpn = _FO_PathSplit($gaDropFiles[0])
					If $aExpn[2] == '.7_' Then
						ShellExecuteWait('C:\Program Files\7-Zip\7z.exe', 'x ' & $gaDropFiles[0] & ' -o' & $aExpn[0] & ' -aoa')
						GUICtrlSetImage($Pic1, "C:\Temp\done.jpg")
					Else
						$i = 0
						FileDelete('C:\Temp\Archive.7_')
						ShellExecuteWait('C:\Program Files\7-Zip\7z.exe', 'a "' & 'C:\Temp\Archive.7_"  -mhe -mx9 "' & $gaDropFiles[0], '', '', @SW_HIDE)
						If FileGetSize('C:\Temp\Archive.7_') < 18*1024*1024 Then
							CreateMailItem()
							GUICtrlSetImage($Pic1, "C:\Temp\done.jpg")
						Else
							MsgBox(4096, 'Error', "Вы не можете отправить этот файл! "& @CRLF &  "Его размер первышает допустимый.")
						EndIf
					EndIf
				Case Else
					FileDelete('C:\Temp\Archive.7_')
					_FileWriteFromArray('C:\Temp\listfile.txt', $gaDropFiles, 0)
					ShellExecuteWait('C:\Program Files\7-Zip\7z.exe', 'a "' & 'C:\Temp\Archive.7_ "' & ' "'& '@C:\Temp\listfile.txt" ' & '-mhe -mx9' , '', '', @SW_HIDE)
					If FileGetSize('C:\Temp\Archive.7_') < 18*1024*1024 Then
						CreateMailItem()
						GUICtrlSetImage($Pic1, "C:\Temp\done.jpg")
					Else
						MsgBox(4096, 'Error', "Вы не можете отправить этот файл! "& @CRLF &  "Его размер первышает допустимый.")
					EndIf
			EndSwitch
		Case FileExists('C:\Program Files (x86)\7-Zip\7z.exe')
			ShellExecuteWait('C:\Program Files (x86)\7-Zip\7z.exe', 'a "' & 'C:\Archive.7z" -p"Мой пароль" -mhe -mx9 "' & 'C:\P1005.log"', '', '', @SW_HIDE)
		Case Else
			MsgBox(4096, '', @ProgramFilesDir & '1')
	EndSelect
EndFunc



Func CreateMailItem()
    Const $olByValue = 1
    Local $olMailItem    = 0
    Local $olFormatRichText = 2
    Local $olImportanceLow   = 0
    Local $olImportanceNormal= 1
    Local $olImportanceHigh  = 2

    $oOApp = ObjCreate("Outlook.Application")
    $oOMail = $oOApp.CreateItem($olMailItem)

    With $oOMail
;~         .To = ("to@domain.com")
;~         .Subject = "email subject"
        .BodyFormat =  $olFormatRichText
        .Importance = $olImportanceNormal
;~         .Body = "email message"
		.HTMLBody = '<HTML><BODY><div><p>&nbsp</p><p>&nbsp</p></div><span><img src="C:\Temp\span.jpg"></span></BODY></HTML>'
        .Attachments.Add ($file, $olByValue , 1)
        .Display
        ;.Send ; doesn't work for security reason
    EndWith
EndFunc