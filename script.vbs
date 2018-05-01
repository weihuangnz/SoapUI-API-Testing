Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Users\DEV.WeiH\Desktop\WebAPI.scripts\Results\End-to-end.xlsx")

Set objWorksheet = objWorkbook.Worksheets(1)


i = 1


Do Until objExcel.Cells(i, 1) = ""

    strValue = objExcel.Cells(i, 18)


    If (strValue = "NO" ) Then

        objExcel.Cells(i, 1).EntireRow.Interior.ColorIndex = 44

    End If


    i = i + 1

Loop

objWorkbook.Save
objWorkbook.Close 
objExcel.Quit
