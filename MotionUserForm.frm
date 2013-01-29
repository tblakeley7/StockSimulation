VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MotionUserForm 
   Caption         =   "Welcome to the Geometric Browian Motion Application!"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15705
   OleObjectBlob   =   "MotionUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MotionUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CancelBTN_Click()
Unload Me
End Sub


Private Sub PickSaveAsLocationBTN_Click()

'Check if the Users have the WordReportCHKBX checked off
If WordReportCHKBX.Value = True Then
    '+========================================================================================================+
    ' Allow User to Pick a Folder to save the word report too
        
       Set fd = Application.FileDialog(msoFileDialogFolderPicker)
         With fd
            'Start the Dialog Box @ ThisWorkbook.Path
            .InitialFileName = ThisWorkbook.Path
            
           If .Show Then
            'MsgBox .SelectedItems(1)
                'Store the Selected folder name into a public variable to be used when saving the late binded Word Object
                SaveAsString = .SelectedItems(1)
                'Store the Selected folder name into the textbox below this button
                TempArray(0) = .SelectedItems(1)
                SaveAsTXTBX.List = TempArray
           End If
         End With
    '+========================================================================================================+

Else
    'Force the Checkbox
    WordReportCHKBX.Value = True
    'Then Run Code
    '+========================================================================================================+
    ' Allow User to Pick a Folder to save the word report too
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
     With fd
        'Start the Dialog Box @ ThisWorkbook.Path
        .InitialFileName = ThisWorkbook.Path
        
       If .Show Then
        'MsgBox .SelectedItems(1)
            'Store the Selected folder name into a public variable to be used when saving the late binded Word Object
            SaveAsString = .SelectedItems(1)
            'Store the Selected folder name into the textbox below this button
            TempArray(0) = .SelectedItems(1)
            SaveAsTXTBX.List = TempArray
       End If
     End With
     '+========================================================================================================+
End If

End Sub


Private Sub RunButton_Click()
Worksheets("FrontEnd").Unprotect

Dim numberofxvalues As Integer
Dim xaxisvalues() As Variant
Dim time As Double
Dim count As Integer
Dim I As Integer
Dim xarray() As Variant
Dim stock As Variant
Dim drift As Variant
Dim uncertainty As Variant
Dim stockprice() As Double
Dim j As Long
Dim m As Long
Dim K As Long
Dim wks As Worksheet
Dim FinalStockPrices() As Variant
Dim newruncounter As Long
Dim celllcounter As Long
Dim Cell As Range
Dim ctl As Control

'Delete the BrownianMotion worksheet if it already exists before the run
For Each wks In Application.ActiveWorkbook.Worksheets
    If wks.Name = "BrownianMotion" Then
        Application.DisplayAlerts = False
        wks.Delete
        Application.DisplayAlerts = True
    End If
Next

'Allocate the User Selected Stocks 3-Char ID and Price
For K = 0 To StockPriceLstBX.ListCount - 1
  If StockPriceLstBX.Selected(K) = True Then
     Data = StockPriceLstBX.List(K, 1)
  End If
Next K

For K = 0 To StockPriceLstBX.ListCount - 1
    If StockPriceLstBX.Selected(K) = True Then
    StockName = StockPriceLstBX.List(K)
    End If
Next K

'Form Error Handling

If IsNumeric(Data) = False Then
    MsgBox "Please do not select the header in the list box", vbCritical, "List Box Error"
    Exit Sub
End If

For Each ctl In Me.Controls
    If TypeName(ctl) = "TextBox" Then
        If ctl.Value = "" Or IsNumeric(DriftTXTBX.Value) = False Or NTXTBX.Value > 255 Or NTXTBX.Value < 0 Then

            If VolatilityTXTBX.Value = "" Or IsNumeric(VolatilityTXTBX.Value) = False Then
                MsgBox "Please Enter a Volatility that is Numeric and between 0 and 1.", vbCritical, "Volatility Input Error"
                VolatilityTXTBX.SetFocus
            ElseIf DriftTXTBX.Value = "" Or IsNumeric(DriftTXTBX.Value) = False Then
                MsgBox "Please Enter a Drift that is Numeric and between 0 and 1.", vbCritical, "Drift Input Error"
                DriftTXTBX.SetFocus
            ElseIf TimeTXTBX.Value = "" Or IsNumeric(TimeTXTBX.Value) = False Then
                MsgBox "Please Enter a Time Interval (0.5 makes each node a half year in length from each other) ", vbCritical, "Time Interval Input Error"
                TimeTXTBX.SetFocus
            ElseIf TUBoundTXTBX.Value = "" Or IsNumeric(TUBoundTXTBX.Value) = False Then
                MsgBox "Please Enter a Time Ubound that will complement 'n' runs of GBM (Between 0 and 2 [years]) ", vbCritical, "Time Upper Bound Input Error"
                TUBoundTXTBX.SetFocus
            ElseIf NTXTBX.Value = "" Or IsNumeric(NTXTBX.Value) = False Or NTXTBX.Value > 255 Or NTXTBX.Value < 0 Then
                MsgBox "Please Enter the number of times you want to run Geometric Brownian Motion. There is a minimum of 0 and a limit of 256.", vbCritical, "Volatility Input Error"
                NTXTBX.SetFocus
            End If
            Exit Sub
        End If
    End If
Next

'Set Index to which line the user picked in the Stocks Data List Box
Index = StockPriceLstBX.ListIndex

'Set a counter for the number of times the GBM model is run
newruncounter = 0
For m = 1 To CInt(NTXTBX.Value)

        'Get the Number of time nodes
        numberofxvalues = CDbl(TUBoundTXTBX.Value) / CDbl(TimeTXTBX)
        
        'Size the stockprice array to the number of x values
        ReDim stockprice(numberofxvalues - 1)
        'Set Stock equal to the selected stocks price (Data variable)
        stock = Data
        
        For j = 0 To numberofxvalues - 1
        
            'Run Geometric Brownian Motion
            drift = TimeTXTBX * DriftTXTBX * stock
            uncertainty = WorksheetFunction.NormSInv(Rnd()) * Sqr(CDbl(TimeTXTBX)) * CDbl(VolatilityTXTBX) * stock
            'Add the three for the new stock price
            stockprice(j) = drift + uncertainty + stock
            stock = stockprice(j)
            
            'Dump the last stock price for the nth-run into a FinalStockPrices array to average and count for a trading strategy
            If j = (numberofxvalues - 1) Then
                ReDim Preserve FinalStockPrices(newruncounter)
                FinalStockPrices(newruncounter) = stock 'stockprice(j)
                newruncounter = newruncounter + 1
            End If
            
        Next
        
        
        ReDim xaxisvalues(numberofxvalues - 1)
        
        Do Until count >= numberofxvalues
            time = time + CDbl(TimeTXTBX)
        
            xaxisvalues(count) = time
            count = count + 1
        Loop
        
        If m = 1 Then
        Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = "BrownianMotion"
        Worksheets("BrownianMotion").Shapes.AddChart.Select
        ActiveChart.ChartType = xlLine
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(1).xvalues = xaxisvalues
            ActiveChart.SeriesCollection(m).Values = stockprice
        Else
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(m).Values = stockprice
        End If

Next

celllcounter = 0
For Each Cell In Worksheets("BrownianMotion").Range("B2", Range("B2").End(xlDown))
    If celllcounter = CInt(NTXTBX.Value) Then
        Exit For
    End If
    Cell.Value = FinalStockPrices(celllcounter)
    celllcounter = celllcounter + 1
Next

With Worksheets("BrownianMotion")
    .Range("B1").Value = "Final Stock Price For Each Run of the GBM Model for: " & StockName
        .Range("B1").End(xlDown).Offset(1, 0).Select
        With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlDouble
                .Weight = xlThick
            End With
        End With
    .Range("B1").End(xlDown).Offset(1, 0).Formula = "=AVERAGE(" & .Range("B2", .Range("B2").End(xlDown)).Address & ")"
    RunAverage = CCur(WorksheetFunction.Average(.Range("B2", .Range("B2").End(xlDown))))
    .Columns("A:D").EntireColumn.AutoFit
    
End With

Call ShowUserFormMacro.Query2


End Sub

Private Sub UserForm_Initialize()

'Populate Stock Data List Box from Database with initial stock prices and stock 3-char IDs
Call ShowUserFormMacro.Query1
'Select first Stock in the list
StockPriceLstBX.ListIndex = 1
'Testing User Inputs
'VolatilityTXTBX.Value = "0.9"
'TimeTXTBX = "0.1"
'DriftTXTBX = "0.8"
'TUBoundTXTBX = "1"
'NTXTBX = "10"

End Sub

