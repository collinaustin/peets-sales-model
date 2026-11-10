Option Explicit

' ============================================================
' OrderEntry > SalesData (tblSales)
'
' Copies one ticket from the OrderEntry tab into tblSales.
' It writes the 8 fields a ticket records:
'   Ticket ID, Visit Date, Store #, SKU, Units, Unit Price,
'   Discount per Unit, Display Units
' It also writes 6 lookup fields as plain values, so tblSales
' doesn't need hundreds of thousands of live XLOOKUP formulas:
'   Unit Cost = Unit Price x (1 - Gross Margin %) from tblSKU
'   Brand, Format, Type, Seasonal from tblSKU
'   Week from tblCal
' Segment stays a live formula, because the Calculated Segment
' in StoreMaster changes as new sales come in.
' Every other column is a calculated column, so Excel fills
' it in automatically for each new row.
'
' OrderEntry layout this macro expects:
'   B4 = Visit Date, B5 = Store #
'   Line rows 11 to 30:
'     A = SKU, C = Units, D = Unit Price (formula), E = Discount per Unit (formula),
'     G = "On promo" flag (formula)
'     I = Display Units (optional: units from this line that went onto a display)
' ============================================================

Private Const ENTRY_SHEET As String = "OrderEntry"
Private Const FIRST_LINE As Long = 11
Private Const LAST_LINE As Long = 30

Public Sub SubmitTicket()
    Dim wsE As Worksheet
    Dim loSales As ListObject, loSKU As ListObject, loStore As ListObject, loCalendar As ListObject
    Dim visitDate As Date, storeNum As Long, ticketID As String
    Dim r As Long, nLines As Long, nPromoNoDisc As Long
    Dim sku As String, units As Variant, price As Variant, disc As Variant
    Dim lr As ListRow
    Dim colID As Long, colDate As Long, colStore As Long, colSKU As Long
    Dim colUnits As Long, colPrice As Long, colDisc As Long
    Dim colCost As Long, colBrand As Long, colFormat As Long, colType As Long
    Dim colSeasonal As Long, colWeek As Long
    Dim weekStart As Date, calRow As Variant, weekNum As Variant
    Dim rngSKU As Range, skuRow As Variant, margin As Variant
    Dim colDisplay As Long, dispUnits As Variant
    Dim prevCalc As Long
    Dim succeeded As Boolean

    Set wsE = ThisWorkbook.Worksheets(ENTRY_SHEET)
    Set loSales = FindTable("tblSales")
    Set loSKU = FindTable("tblSKU")
    Set loStore = FindTable("tblStore")
    Set loCalendar = FindTable("tblCal")

    If loSales Is Nothing Or loSKU Is Nothing Or loStore Is Nothing Or loCalendar Is Nothing Then
        MsgBox "Could not find tblSales, tblSKU, tblStore, or tblCal." & vbCr & vbCr & "Tables in this workbook:" & vbCr & ListAllTables(), vbExclamation
        Exit Sub
    End If

    ' ---------- 1. Check the ticket header ----------
    If Not IsDate(wsE.Range("B4").Value) Then
        MsgBox "Enter a Visit Date in B4.", vbExclamation
        Exit Sub
    End If
    visitDate = CDate(wsE.Range("B4").Value)

    If IsEmpty(wsE.Range("B5").Value) Or Not IsNumeric(wsE.Range("B5").Value) Then
        MsgBox "Enter a Store # in B5.", vbExclamation
        Exit Sub
    End If
    storeNum = CLng(wsE.Range("B5").Value)

    If Application.WorksheetFunction.CountIf(loStore.ListColumns("Store #").DataBodyRange, storeNum) = 0 Then
        MsgBox "Store # " & storeNum & " is not in StoreMaster.", vbExclamation
        Exit Sub
    End If

    ticketID = "T" & Format(visitDate, "yymmdd") & "-" & storeNum
    If Application.WorksheetFunction.CountIf(loSales.ListColumns("Ticket ID").DataBodyRange, ticketID) > 0 Then
        MsgBox "Ticket " & ticketID & " is already in SalesData." & vbCr & "Nothing was added.", vbExclamation
        Exit Sub
    End If

    ' The ticket's week must be in tblCal so the Week column can be filled in.
    ' Same math as the Week Start column: Visit Date - WEEKDAY(Visit Date, 3), so weeks start Monday.
    weekStart = visitDate - (Weekday(visitDate, vbMonday) - 1)
    calRow = Application.Match(CLng(weekStart), loCalendar.ListColumns("Week Start").DataBodyRange, 0)
    If IsError(calRow) Then
        MsgBox "The week starting " & Format(weekStart, "m/d/yyyy") & " is not on the Calendar tab." & vbCr & "Add that week to tblCal first. Nothing was added.", vbExclamation
        Exit Sub
    End If
    weekNum = loCalendar.ListColumns("Week").DataBodyRange.Cells(calRow, 1).Value

    ' ---------- 2. Check every line before writing anything ----------
    For r = FIRST_LINE To LAST_LINE
        sku = Trim(CStr(wsE.Cells(r, 1).Value))
        If sku <> "" Then
            If Application.WorksheetFunction.CountIf(loSKU.ListColumns("SKU").DataBodyRange, sku) = 0 Then
                MsgBox "Row " & r & ": SKU " & sku & " is not in SKUMaster.", vbExclamation
                Exit Sub
            End If

            units = wsE.Cells(r, 3).Value
            If IsEmpty(units) Or Not IsNumeric(units) Then
                MsgBox "Row " & r & ": enter Units for " & sku & ".", vbExclamation
                Exit Sub
            End If
            If units <= 0 Or units <> Int(units) Then
                MsgBox "Row " & r & ": Units must be a whole number above 0.", vbExclamation
                Exit Sub
            End If

            price = wsE.Cells(r, 4).Value
            If IsError(price) Then
                MsgBox "Row " & r & ": Unit Price shows an error.", vbExclamation
                Exit Sub
            End If
            If IsEmpty(price) Or Not IsNumeric(price) Then
                MsgBox "Row " & r & ": Unit Price is missing.", vbExclamation
                Exit Sub
            End If

            disc = wsE.Cells(r, 5).Value
            If IsEmpty(disc) Then disc = 0
            If Not IsNumeric(disc) Then
                MsgBox "Row " & r & ": Discount must be a number (or blank for none).", vbExclamation
                Exit Sub
            End If
            If disc < 0 Or disc >= price Then
                MsgBox "Row " & r & ": Discount must be 0 or more and less than the Unit Price.", vbExclamation
                Exit Sub
            End If

            dispUnits = wsE.Cells(r, 9).Value
            If IsEmpty(dispUnits) Then dispUnits = 0
            If Not IsNumeric(dispUnits) Then
                MsgBox "Row " & r & ": Display Units must be a number (or blank for none).", vbExclamation
                Exit Sub
            End If
            If dispUnits < 0 Or dispUnits > units Or dispUnits <> Int(dispUnits) Then
                MsgBox "Row " & r & ": Display Units must be a whole number from 0 up to the line's Units.", vbExclamation
                Exit Sub
            End If

            If Not IsError(wsE.Cells(r, 7).Value) Then
                If wsE.Cells(r, 7).Value = "On promo" And disc = 0 Then
                    nPromoNoDisc = nPromoNoDisc + 1
                End If
            End If

            nLines = nLines + 1
        End If
    Next r

    If nLines = 0 Then
        MsgBox "There are no SKU lines to submit.", vbExclamation
        Exit Sub
    End If

    If nPromoNoDisc > 0 Then
        If MsgBox(nPromoNoDisc & " line(s) are on promo this week but have no discount." & vbCr & "Submit anyway?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    If MsgBox("Add ticket " & ticketID & " (" & nLines & " lines) to SalesData?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ' ---------- 3. Write the lines to tblSales ----------
    colID = loSales.ListColumns("Ticket ID").Index
    colDate = loSales.ListColumns("Visit Date").Index
    colStore = loSales.ListColumns("Store #").Index
    colSKU = loSales.ListColumns("SKU").Index
    colUnits = loSales.ListColumns("Units").Index
    colPrice = loSales.ListColumns("Unit Price").Index
    colDisc = loSales.ListColumns("Discount per Unit").Index
    colCost = loSales.ListColumns("Unit Cost").Index
    colBrand = loSales.ListColumns("Brand").Index
    colFormat = loSales.ListColumns("Format").Index
    colType = loSales.ListColumns("Type").Index
    colSeasonal = loSales.ListColumns("Seasonal").Index
    colWeek = loSales.ListColumns("Week").Index
    colDisplay = loSales.ListColumns("Display Units").Index
    Set rngSKU = loSKU.ListColumns("SKU").DataBodyRange

    ' Remember the calculation mode so it can be put back afterward.
    ' (The old version always switched Excel to Automatic at the end.)
    prevCalc = Application.Calculation

    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For r = FIRST_LINE To LAST_LINE
        sku = Trim(CStr(wsE.Cells(r, 1).Value))
        If sku <> "" Then
            price = wsE.Cells(r, 4).Value
            disc = wsE.Cells(r, 5).Value
            If IsEmpty(disc) Then disc = 0
            dispUnits = wsE.Cells(r, 9).Value
            If IsEmpty(dispUnits) Then dispUnits = 0

            ' Find this SKU's row in tblSKU once, then read every field from it.
            skuRow = Application.Match(sku, rngSKU, 0)
            margin = loSKU.ListColumns("Gross Margin %").DataBodyRange.Cells(skuRow, 1).Value

            Set lr = loSales.ListRows.Add
            ' Ticket fields
            lr.Range.Cells(1, colID).Value = ticketID
            lr.Range.Cells(1, colDate).Value = visitDate
            lr.Range.Cells(1, colStore).Value = storeNum
            lr.Range.Cells(1, colSKU).Value = sku
            lr.Range.Cells(1, colUnits).Value = wsE.Cells(r, 3).Value
            lr.Range.Cells(1, colPrice).Value = price
            lr.Range.Cells(1, colDisc).Value = disc
            lr.Range.Cells(1, colDisplay).Value = dispUnits
            ' Lookup fields, written as plain values
            lr.Range.Cells(1, colCost).Value = price * (1 - margin)
            lr.Range.Cells(1, colBrand).Value = loSKU.ListColumns("Brand").DataBodyRange.Cells(skuRow, 1).Value
            lr.Range.Cells(1, colFormat).Value = loSKU.ListColumns("Format").DataBodyRange.Cells(skuRow, 1).Value
            lr.Range.Cells(1, colType).Value = loSKU.ListColumns("Type").DataBodyRange.Cells(skuRow, 1).Value
            lr.Range.Cells(1, colSeasonal).Value = loSKU.ListColumns("Seasonal").DataBodyRange.Cells(skuRow, 1).Value
            lr.Range.Cells(1, colWeek).Value = weekNum

            ' Calculate just this new row (Net Price, Line Total, Segment, etc.)
            lr.Range.Calculate
        End If
    Next r

    ' ---------- 4. Clear the form for the next store ----------
    ' Visit Date stays, since the next ticket is usually the same day.
    wsE.Range("B5").ClearContents
    wsE.Range("A" & FIRST_LINE & ":A" & LAST_LINE).ClearContents
    wsE.Range("C" & FIRST_LINE & ":C" & LAST_LINE).ClearContents
    wsE.Range("I" & FIRST_LINE & ":I" & LAST_LINE).ClearContents
    succeeded = True

CleanUp:
    If prevCalc <> 0 Then Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    If succeeded Then
        If prevCalc = xlCalculationManual Then
            MsgBox "Ticket " & ticketID & " added: " & nLines & " lines." & vbCr & "Calculation is on Manual, so press F9 to update the reports.", vbInformation
        Else
            MsgBox "Ticket " & ticketID & " added: " & nLines & " lines.", vbInformation
        End If
    Else
        MsgBox "Something went wrong: " & Err.Description & vbCr & "Check the bottom of SalesData for a partial ticket.", vbCritical
    End If
End Sub

' ============================================================
' ONE-TIME SETUP: run this once, after installing this module.
' Recalculates everything, checks for errors, then replaces the
' formulas in 6 lookup columns of tblSales with their values:
'   Unit Cost, Brand, Format, Type, Seasonal, Week
' Segment is left as a formula on purpose.
' ============================================================
Public Sub FreezeLookupColumns()
    Dim loSales As ListObject, rng As Range, errCells As Range
    Dim colNames As Variant, i As Long, prevCalc As Long

    Set loSales = FindTable("tblSales")
    If loSales Is Nothing Then
        MsgBox "Could not find tblSales.", vbExclamation
        Exit Sub
    End If

    If MsgBox("This replaces the formulas in Unit Cost, Brand, Format, Type, Seasonal, and Week with their current values." & vbCr & _
              "It can't be undone, so save a backup copy first." & vbCr & vbCr & "Continue?", vbYesNo + vbExclamation) = vbNo Then Exit Sub

    colNames = Array("Unit Cost", "Brand", "Format", "Type", "Seasonal", "Week")
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False

    ' Make sure every value is current before freezing it.
    ' This is the slow part: expect the usual recalculation freeze.
    Application.CalculateFull

    ' Stop if any lookup shows an error, so no error gets frozen in place.
    For i = LBound(colNames) To UBound(colNames)
        Set rng = loSales.ListColumns(colNames(i)).DataBodyRange
        Set errCells = Nothing
        On Error Resume Next
        Set errCells = rng.SpecialCells(xlCellTypeFormulas, xlErrors)
        On Error GoTo 0
        If Not errCells Is Nothing Then
            Application.ScreenUpdating = True
            MsgBox colNames(i) & " has " & errCells.Count & " error cell(s), first at " & _
                   errCells.Cells(1, 1).Address(False, False) & "." & vbCr & "Nothing was changed. Fix those rows first.", vbExclamation
            Exit Sub
        End If
    Next i

    Application.Calculation = xlCalculationManual
    For i = LBound(colNames) To UBound(colNames)
        Set rng = loSales.ListColumns(colNames(i)).DataBodyRange
        rng.Value = rng.Value
    Next i
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    MsgBox "Done. Unit Cost, Brand, Format, Type, Seasonal, and Week in tblSales are now plain values." & vbCr & "Save the workbook to keep the change.", vbInformation
End Sub

' Finds a table by name on any tab, so tab names don't matter.
Private Function FindTable(tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase(lo.Name) = LCase(tableName) Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' Lists every table and the tab it is on, for the error message.
Private Function ListAllTables() As String
    Dim ws As Worksheet, lo As ListObject, s As String
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            s = s & lo.Name & "  (tab: " & ws.Name & ")" & vbCr
        Next lo
    Next ws
    If s = "" Then s = "(none found)"
    ListAllTables = s
End Function
