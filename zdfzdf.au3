Global $file = "C:\Temp\Archive.7_"
CreateMailItem()
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