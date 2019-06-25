;Start;
saveFile = PullSlips.docx
FileDelete %saveFile%

;Configurations
Iniread, download_directory, config.ini, general, download_directory
Iniread, version, config.ini, general, version

;Check for Template.docx
IfNotExist, Template.docx
{
	msgbox Cannot find Template.docx
	exit
}

;Get input file
FileSelectFile, xlsFile,,C:\Users\%A_UserName%\Downloads\, PullSlips %version% - Select File, *.xls*

;Check for input file or cancel to exit
If xlsFile =
{
	exit
}

;Open XLS file
xl := ComObjCreate("Excel.Application")
xl.Visible := False
book := xl.Workbooks.Open(xlsFile)
rows := book.Sheets(1).UsedRange.Rows.Count

;Sort by HeldFor Column
xlAscending := 1
xlYes := 1
book.Sheets(1).UsedRange.Sort(Key1 := xl.Range("I2")
		, Order1 := xlAscending,,,,,
		, Header := xlYes)
        
;Save and quit XLS file
book.Save()
book.Close
xl.Quit

;Progress
Progress, zh0 fs12, Generating Slips...,,PullSlips

;Open DOC file
template = %A_ScriptDir%\Template.docx
saveFilePath = %A_ScriptDir%\%saveFile%
wrd := ComObjCreate("Word.Application")
wrd.Visible := False

;Perform Mail Merge
doc := wrd.Documents.Open(template)
doc.MailMerge.OpenDataSource(xlsFile,,,,,,,,,,,,,"SELECT * FROM [expiredHoldShelfRequestsList$] WHERE [Location] = 'ILLIAD' OR [Location] = 'Resource Sharing Long Loan' OR [Location] = 'Resource Sharing Short Loan'")
doc.MailMerge.Execute

;Save and quit DOC file
wrd.ActiveDocument.SaveAs(saveFilePath)
wrd.DisplayAlerts := False
doc.Close
wrd.Quit

;Progress
Progress, zh0 fs12, Sending to Word...,,PullSlips

;Finish
IfNotExist, %saveFile%
{
	msgbox Cannot find %saveFile%
	exit
}
;FileDelete %xlsFile%
FileDelete %xlsFile%
run winword.exe %saveFile%