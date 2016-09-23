Dim excelApp

Set excelApp = CreateObject("Excel.Application")

	excelApp.Application.Visible = False
	excelApp.Application.Run "Tracking sheet V4.xlsm!call_USFX"
	WScript.Echo "Done"