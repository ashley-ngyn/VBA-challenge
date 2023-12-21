# Module 2

this readme shows screenshots of excel next to vba code and all sources used


here are screenshots of my vba code next to excel:

below shows the entire code ran next to the first excel sheet labeled "2018"
<img width="1470" alt="1" src="https://github.com/ashley-ngyn/module_2/assets/150317761/adc10139-cdcf-4596-86b0-c78ce1f7ac38">
<img width="1470" alt="2" src="https://github.com/ashley-ngyn/module_2/assets/150317761/39b82e8d-b653-41d5-9c17-2f185169738a">
<img width="1470" alt="3" src="https://github.com/ashley-ngyn/module_2/assets/150317761/b9b758a1-e7e7-4781-8792-bb115d052fc3">
<img width="1470" alt="4" src="https://github.com/ashley-ngyn/module_2/assets/150317761/fe8b41a9-8d91-4382-82ba-6fc9715b2ecb">

this shows the second excel sheet labeled "2019"
<img width="1470" alt="5" src="https://github.com/ashley-ngyn/module_2/assets/150317761/7a2f4fab-1197-417a-b17b-56eb5a4e2bcb">

this shows the third excel sheet labeled "2020"
<img width="1470" alt="6" src="https://github.com/ashley-ngyn/module_2/assets/150317761/1e14e139-93c2-4c0d-ad33-9a88a9a9701f">


sources used:

  used for applying code to all worksheets

    For Each ws In Worksheets
    
    https://support.microsoft.com/en-au/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
    
    https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop

  used to count all the rows after the column selected

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lastrowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    https://www.excelcampus.com/vba/find-last-row-column-cell/

  format the numbers to percent or scientific notation

    ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).NumberFormat = "0.00E+0"
   
    https://www.automateexcel.com/vba/format-numbers/

  finding the max or min

    ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowSummary))
    ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowSummary))
    ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowSummary))
   
    https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid=79&Itemid=475

  class activities
    
    credit card charges (for loop)
    For i = 2 To lastrow
    
    formatter (color changes)
    ws.Cells(i, 10).Interior.ColorIndex = 4
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
    grader solution (conditionals with color changes)
    If ws.Cells(i, 10).Value > 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 4
