# realignment
Sub mk_template()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Tdata = ActiveWorkbook.Name

Dim sh As Worksheet
Dim LastRowC As Long
Dim LastRowY As Long
Dim FR1 As Long

'Update File Path
FPath = "C:\Documents and Settings\SENAGAPALLI.KARTHIK\My Documents\Excel Forum\Gti182"
'Set sh = Workbooks("" & Tdata & "").Sheets("Sheet1")

With Workbooks("Test_Data.xlsm").Sheets("Sheet1")
    
    LastRowY = .Range("Y" & .Rows.Count).End(xlUp).Row
    
        'This opens up the Journal template so its ready to for data transfer
        Workbooks.Open FPath & "\" & "Test_Journal.xls"
        
        'This sets Variable "Tjournal" to Test_Journal.xls
        Tjounral = ActiveWorkbook.Name
        
        
     LastRowC = Workbooks("Test_Journal.xls").Sheets("Entry").Range("C" & .Rows.Count).End(xlUp).Row
     
     Stat1 = LastRowY - 3
     Stat2 = LastRowC - 17
     
     AddRow = Stat2 - Stat1
      
     If Stat1 > Stat2 Then
     
        MsgBox "Journal Temple requires addtional rows = " & AddRow & "before it can proceed"
        
        Else
                 .Range("Y4:AK" & LastRowY).SpecialCells(xlCellTypeVisible).Copy _
                    Destination:=Workbooks("Test_Journal.xls").Sheets("Entry").Range("C16:O" & LastRowC - 2)
                    
                 'Need to still transfer other side of journal data
                 'Easiest way is to simply copy and paste it at the bottom of the data at the very beginning
                    
                    
     End If
        
 End With
 
 Workbooks("Test_Journal.xls").Sheets("Entry").Activate
 
 With Worksheets("Entry")
 
    Bus = Cells(16, 3).Value
    Mth = MonthName(Month(Date - 1)) & "-" & Year(Date)
    
    
    .Range("H9").Value = Format((Date - 4), "DD/MM/YY")
    .Range("H10").FormulaR1C1 = "Bus " & Bus & " AL Realignment Journal" & Mth
    .Range("H11").FormulaR1C1 = "Bus " & Bus & " AL Realignment Journal" & Mth
    
End With


        ActiveWorkbook.SaveAs Filename:=FPath & "\Bus " & .Range("C16") & "_" & .Range("M16") & "_" & Format(Now(), "mm_dd_yyyy hh mm AMPM") & ".xls", FileFormat:=xlExcel8
        ActiveWorkbook.Close False
        
        
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

