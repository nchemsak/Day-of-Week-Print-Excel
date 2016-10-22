On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_UTCTime",,48)

For Each objItem in colItems
  Set objExcel = CreateObject("Excel.Application")
    If (objItem.DayOfWeek = "1") Then
      Set objWorkbook = objExcel.Workbooks.Open("")
      objWorkbook.PrintOut()
    ElseIf (objItem.DayOfWeek = "2") Then
      Set objWorkbook = objExcel.Workbooks.Open("")
      objWorkbook.PrintOut()
    ElseIf (objItem.DayOfWeek = "3") Then
      Set objWorkbook = objExcel.Workbooks.Open("")
      objWorkbook.PrintOut()
    ElseIf (objItem.DayOfWeek = "4") Then
      Set objWorkbook = objExcel.Workbooks.Open("")
      objWorkbook.PrintOut()
    ElseIf (objItem.DayOfWeek = "5") Then
      Set objWorkbook = objExcel.Workbooks.Open("")
      objWorkbook.PrintOut()
  End  If
Next
objExcel.Quit
