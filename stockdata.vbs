'This is the moderate solution
'This will sum the volumes of each ticker and display them for each year (worksheet)

Sub volumetotal()
'Loop Through Each worksheet
    For Each ws In Worksheets
'Declare variables
    Dim i As Long
    Dim lastrow As Long
    Dim lastrow1 As Long
    Dim volume As Double
    Dim summaryrow As Integer
    Dim ticker As String
    Dim oprice As Double
    Dim cprice As Double
    Dim prow As Double
    Dim firstopen As Double
    Dim change As Double
    Dim j As Integer
    Dim pchange As Double
    Dim pchangemax As Double
    Dim pchangemin As Double
    Dim volmax As Double
    Dim maxticker As String
    Dim minticker As String
    Dim volticker As String
    
'Assign values
    volume = 0
    summaryrow = 2
    oprice = ws.Cells(2, 3).Value
    cprice = 0
    change = 0
    pchange = 0
'Initial open Value
    ws.Cells(2, 10).Value = oprice
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Loop through to sum volumes by ticker
        For i = 2 To lastrow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            ticker = ws.Cells(i, 1)
            volume = volume + ws.Cells(i, 7).Value
            cprice = ws.Cells(i, 6)
            oprice = ws.Cells(i + 1, 3)
            change = cprice - oprice
'Write the Volume of the ticker
            ws.Cells(summaryrow, 9).Value = ticker
            ws.Cells(summaryrow + 1, 10).Value = oprice
            ws.Cells(summaryrow, 11).Value = cprice
            ws.Cells(summaryrow, 12).Value = volume
            
                       
'Update variables
            summaryrow = summaryrow + 1
            volume = 0
            Else
            volume = volume + ws.Cells(i, 7).Value
            
            End If
        Next i
    ws.Cells(summaryrow, 10).Value = ""
  
      
'Calculate change and percent change
    lastrow1 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To lastrow1
            change = ws.Cells(j, 11).Value - ws.Cells(j, 10).Value
                If ws.Cells(j, 10) = 0 Then
                    pchange = 0
                Else: pchange = change / ws.Cells(j, 10).Value
                End If
            ws.Cells(j, 10).Value = change
            ws.Cells(j, 11).Value = pchange
            ws.Cells(j, 11).NumberFormat = "0.00%"
                If (ws.Cells(j, 10)) > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                Else
                     ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
        Next j
'Finding Largest Changes and greatest volume
        For j = 2 To lastrow1
             If ws.Cells(j, 11).Value > pchangemax Then
             pchangemax = ws.Cells(j, 11).Value
             maxticker = ws.Cells(j, 9).Value
             End If
                If ws.Cells(j, 11).Value < pchangemin Then
                pchangemin = ws.Cells(j, 11).Value
                minticker = ws.Cells(j, 9).Value
                End If
                   If ws.Cells(j, 12).Value > volmax Then
                   volmax = ws.Cells(j, 12).Value
                   volticker = ws.Cells(j, 9).Value
                   End If
         Next j
        'MsgBox (pchangemax)
'Writing out results of max and min and greatest volume
    ws.Range("o2").Value = maxticker
    ws.Range("p2").Value = pchangemax
    ws.Range("p2").NumberFormat = "0.00%"
    ws.Range("o3").Value = minticker
    ws.Range("p3").Value = pchangemin
    ws.Range("p3").NumberFormat = "0.00%"
    ws.Range("o4").Value = volticker
    ws.Range("p4").Value = volmax
    ws.Range("p4").NumberFormat = "0"
    
        'MsgBox ("next sheet")
'reset values
maxticker = ""
pchangemax = 0
minticker = ""
pchangemin = 0
volticker = ""
volmax = 0
lastrow1 = 0
    'MsgBox ("next sheet")
Next ws
            
            
End Sub
