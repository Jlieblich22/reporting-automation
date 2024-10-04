'Create ListSalesOrders Macro
Sub ListSalesOrders()
'Set Global Dims
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim lRow As Long
   
'Change Original Export sheet name to Raw Data
    Sheets("Sheet1").Name = "Raw Data"
    
'Add three new sheets
    Set ws = Sheets.Add
    Set ws3 = Sheets.Add
    Set ws2 = Sheets.Add
    ws.Name = "Sheet1"
    ws3.Name = "Sheet2"
    ws2.Name = "Sheet3"
   
'Copy the Raw Data to Sheet1
    With ws
        Dim sourceSheet As Worksheet
        Dim destinationSheet As Worksheet
        Dim sourceRange As Range
        Set sourceSheet = Sheets("Raw Data")
        Set destinationSheet = Sheets("Sheet1")
        Set sourceRange = sourceSheet.Range("A:Q")
        sourceRange.Copy destinationSheet.Range("A1")
    End With

    
'Delete Row if non-Client Coordinator order (ZCR and ZDR are non-CC order types)
    With ws
        .AutoFilterMode = False
        lRow = .Range("B" & .Rows.Count).End(xlUp).Row
        Set Rng = .Range("B1:B" & lRow)
        With Rng
            .AutoFilter Field:=1, Criteria1:="ZCR", VisibleDropDown:=False
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .AutoFilter Field:=1, Criteria1:="ZDR", VisibleDropDown:=False
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With
        .AutoFilterMode = False
    End With

'Delete Row if Fuel Surcharge (100100)
    With ws
        .AutoFilterMode = False
        Row = .Range("G" & .Rows.Count).End(xlUp).Row
        Set Rng2 = .Range("G1:G" & Row)
        With Rng2
            .AutoFilter Field:=1, Criteria1:="100100", VisibleDropDown:=False
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With
        .AutoFilterMode = False
    End With

'Remove fields from Sheet1 to  remove need for Change Layout field removal step during SAP Export
'Fields removed:
'Customer Reference, Material Description, Schedule line number, Order Quantity (Item), Confirmed Quantity (Item),
'Delivery Status Description, Delivery Date, Goods Issue Date, Material Availability Date, Partner Name
Sheets("Sheet1").Activate
Range("C:C,I:I,J:J,K:K,L:L,M:M,N:N,O:O,P:P,Q:Q").Delete
    
'Copy/paste all data, except for Material Number Column, from Sheet1 to Sheet3
    With ws
        Dim sourceSheet2 As Worksheet
        Dim destinationSheet2 As Worksheet
        Dim sourceRange2 As Range
        Set sourceSheet2 = Sheets("Sheet1")
        Set destinationSheet2 = Sheets("Sheet3")
        Set sourceRange2 = sourceSheet2.Range("A:F")
        sourceRange2.Copy destinationSheet2.Range("A1")
    End With

'Remove Dupes by Sales Document on last sheet (Sheet3)
    With ws2
        .AutoFilterMode = False
        Set Rng3 = .Range("A:F")
            With Rng3
                .RemoveDuplicates Columns:=1, Header:=xlYes
            End With
        .AutoFilterMode = False
    End With

'Copy/paste only the Created by Column from Sheet3 to Sheet2
    With ws2
        Dim sourceSheet3 As Worksheet
        Dim destinationSheet3 As Worksheet
        Dim sourceRange3 As Range
        Set sourceSheet3 = Sheets("Sheet3")
        Set destinationSheet3 = Sheets("Sheet2")
        Set sourceRange3 = sourceSheet3.Range("D:D")
        sourceRange3.Copy destinationSheet3.Range("A1")
    End With

'Remove Created by Column Dupes on Sheet2
    With ws3
        .AutoFilterMode = False
        Set Rng4 = .Range("A:A")
            With Rng4
                .RemoveDuplicates Columns:=1, Header:=xlYes
            End With
        .AutoFilterMode = False
    End With
    
'Remove SAP_WFRT Created by data point
    With ws3
        Dim Rng5 As Range
        Dim cell As Range
        Dim i As Long
        Set Rng5 = .Range("A:A")
        For i = Rng5.Rows.Count To 2 Step -1
            If Rng5.Cells(i, 1).Value = "SAP_WFRT" Then
                Rng5.Cells(i, 1).EntireRow.Delete
            End If
        Next i
    End With
    
'Add SO Entered - Line Items, SO Entered, and Orders per Day columns and corresponding formulas to Sheet2
    With ws3
        Sheets("Sheet2").Activate
        Cells(1, 2).Value = "SO Entered - Line Items"
        Cells(2, 2).Value = "=COUNTIFS(Sheet1!D:D,A2)"
        Range("B2").AutoFill Destination:=Range("B2:B" & Cells(Rows.Count, "A").End(xlUp).Row)
        Cells(1, 3).Value = "SO Entered"
        Cells(2, 3).Value = "=COUNTIFS(Sheet3!D:D,A2)"
        Range("C2").AutoFill Destination:=Range("C2:C" & Cells(Rows.Count, "A").End(xlUp).Row)
        Cells(1, 4).Value = "Orders per Day"
        Cells(2, 4).Value = "=C2/COUNT(UNIQUE(Sheet3!C:C),A2)"
        Range("D2").AutoFill Destination:=Range("D2:D" & Cells(Rows.Count, "A").End(xlUp).Row)
    End With
End Sub
