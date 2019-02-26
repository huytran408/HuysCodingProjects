Attribute VB_Name = "hard"
Option Explicit

Sub hard()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets

ws.Activate

Dim i As Long
Dim lastrow As Long
Dim ticker As String
Dim tickertotal As Double
Dim openprice As Currency
Dim closeprice As Currency
Dim yearlychange As Currency
Dim Percentchange As Single
Dim counter As Double

Dim counteropen As Double



counteropen = 0
counter = 0
tickertotal = 0

Dim tickertotalrow As Double
tickertotalrow = 2

Range("J1").Value = Range("A1")
Range("K1") = "Yearly Change"
Range("L1") = "Percent Change"
Range("M1") = "Total Volume"
Range("k:K").EntireColumn.NumberFormat = "$#,##0.00"
Range("l:l").EntireColumn.NumberFormat = "0.00%"
Range("k:K").ColumnWidth = 15
Range("l:l").ColumnWidth = 15
Range("m:m").EntireColumn.NumberFormat = "#,##0"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow



    If Cells(i + 1, 1) <> Cells(i, 1) Then
       
        ticker = Cells(i, 1).Value
        tickertotal = tickertotal + Cells(i, 7).Value
        
        openprice = Cells(i - counter, 3).Value
            If openprice = 0 Then
            openprice = Cells(i - (counter - counteropen), 3).Value
            
            End If
        
        closeprice = Cells(i, 6).Value
        yearlychange = closeprice - openprice
        
        If openprice = 0 Then
            Percentchange = 0
        Else
            Percentchange = (closeprice - openprice) / openprice
        End If
        
        Cells(tickertotalrow, 10).Value = ticker
        Cells(tickertotalrow, 11).Value = yearlychange
        Cells(tickertotalrow, 12).Value = Percentchange
        Cells(tickertotalrow, 13).Value = tickertotal
        
        tickertotalrow = tickertotalrow + 1
        tickertotal = 0
        yearlychange = 0
        Percentchange = 0
        counter = 0
        openprice = 0
        closeprice = 0
        
        

    Else
    
    tickertotal = tickertotal + Cells(i, 7)
    counter = counter + 1
        If Cells(i, 3) = 0 Then
        counteropen = counteropen + 1
    
         End If
     
    End If
    
    
Next i

Dim j As Double
Dim lastrowb As Double
Dim max As Double
Dim min As Double
Dim minticker As String
Dim totalvalue As Double
Dim maxticker As String
Dim totalticker As String




lastrowb = Cells(Rows.Count, 10).End(xlUp).Row

Range("p2") = "Greatest % Increase"
Range("p3") = "Greatest % Decrease"
Range("p4") = "Greatest Total Volume"
Range("Q1") = "Ticker"
Range("R1") = "Value"
Range("p:p").ColumnWidth = 20

max = -100000
min = 100000
totalvalue = 0

For j = 2 To lastrowb

    If Cells(j, 11).Value > 0 Then
    Cells(j, 11).Interior.ColorIndex = 4
    ElseIf Cells(j, 11).Value < 0 Then
    Cells(j, 11).Interior.ColorIndex = 3
    
    End If
Next j

    
For j = 2 To lastrowb
    If Cells(j, 12) > max Then
        maxticker = Cells(j, 10).Value
        max = Cells(j, 12).Value
   
    
    End If
    
    
Next j

Range("Q2") = maxticker
Range("R2") = max
Range("R2").NumberFormat = "0.00%"

For j = 2 To lastrowb
    If Cells(j, 12) < min Then
        minticker = Cells(j, 10).Value
        min = Cells(j, 12).Value
   
    
    End If
    
    
Next j

Range("Q3") = minticker
Range("R3") = min
Range("R3").NumberFormat = "0.00%"

For j = 2 To lastrowb
    If Cells(j, 13) > totalvalue Then
        totalticker = Cells(j, 10).Value
        totalvalue = Cells(j, 13).Value
   
    
    End If
    
    
Next j
Range("Q4") = totalticker
Range("R4") = totalvalue
Range("R4").NumberFormat = "#,##0"






Debug.Print ws.Name

Next ws


End Sub


