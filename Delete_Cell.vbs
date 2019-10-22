Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set xlVbscript = objExcel.WorkBooks.Open("C:\Users\Excel.xlsx")

xlVbscript.Sheets(1).Range("A1").Delete'''''''''''Delete a Particular Cell''''''''

xlVbscript.Sheets(1).Range("A1:A5").Delete''''''''''''Delete From A1 To A5''''''''''''

xlVbscript.Sheets(1).Rows(1).EntireRow.Delete'''''''''''Delete Entire 1st Row''''''''''''

xlVbscript.Sheets(1).Columns(3).EntireColumn.Delete'''''''''''Delete Entire 1st Column''''''''''''

xlVbscript.save
xlVbscript.Close

objExcel.Quit
set objExcel=nothing