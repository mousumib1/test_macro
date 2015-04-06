'*****************************************************
' Import macro into given parameter sheet and then run it. Log error if any.
' Pre-requisite --- To add macro in workbook, macro must be enabled and 'Trust access to the VBA project object model' must be checked in given workbook/parameter sheet.
'*****************************************************
Dim oXL 
Dim oBook 
Dim oSheet 
Dim i, J
Dim sMsg
Dim fso, MyFile, FileName
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
'Opens the file using the system default. / Opens the file as Unicode./ Opens the file as ASCII.

On Error resume next 

If WScript.Arguments.Count < 6 Then
    WScript.Echo "Missing parameters. Please enter all five parameter: Log file, Parameter sheet,Sheet name for which macro should work, Macro file, Macro function name. Give full file path for all files."
Else
	Set fso = CreateObject("Scripting.FileSystemObject")

	' Open the file for output.
	FileName = WScript.Arguments(0) ' "F:\aws_work\awscoe\APRIL2015\poc\macro_attach\test.log"
	Set MyFile = fso.OpenTextFile(FileName, ForAppending, True, TristateTrue)
	
	' Create a new instance of Excel and make it visible.
	Set oXL = CreateObject("Excel.Application")
	oXL.Visible = True

	' Open workbook and set a reference to desire sheet.
	Set oBook = oXL.Workbooks.Open(WScript.Arguments(1)) ' "F:\aws_work\awscoe\APRIL2015\poc\macro_attach\IAM_parameter sheet.xlsm"
	Set oSheet = oBook.Sheets(WScript.Arguments(2))' "1_IAM"
	oSheet.Activate

	' The Import method lets you add modules to VBA at
	' run time. Change the file path to match the location
	' of the macro file you created .
	oXL.VBE.ActiveVBProject.VBComponents.Import WScript.Arguments(3) ' "F:\aws_work\awscoe\APRIL2015\poc\macro_attach\KbTest.bas"
	' Now run the macro
	oXL.Run WScript.Arguments(4), WScript.Arguments(5)
	
	' save macro in workbook and close it
	oXL.UserControl = False
	oXL.ActiveWorkbook.Save
	oXL.ActiveWorkbook.Close
	
	' release any outstanding object references.
	Set oSheet = Nothing
	Set oBook = Nothing
	Set oXL = Nothing
	MyFile.Close
	' on error write to file and clear error
	If Err.Number <> 0 Then 
		MyFile.WriteLine Err.Number, Err.Description, Err.Source
		MyFile.Close
		Set oSheet = Nothing
		Set oXL = Nothing
		Err.Clear
	End If
End If










   