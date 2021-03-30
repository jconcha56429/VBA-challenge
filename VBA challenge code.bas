Attribute VB_Name = "Module1"
Sub hw2()
For Each ws In Worksheets
    ' Dimensioning variables
    Dim Ticker As String
    Dim inputrow As Double
    Dim outputrow As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim opentoclose As Double
    Dim percentchange As Double
    Dim volume As Double
    Dim lastrow As Double
    
    outputrow = 2 ' second row where variables will be placed
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row ' grabbing last row
    ' ws.cell(15, 15).Value = lastrow
    openprice = ws.Cells(2, 3).Value ' get the open price for the very first stock
    volume = 0
     For inputrow = 2 To lastrow 'grabbing rows
        If ws.Cells(inputrow + 1, 1).Value <> ws.Cells(inputrow, 1).Value Then ' if next row is not = to current row
            ' Debug.Print "openprice", openprice
            ' Inputs
            Ticker = ws.Cells(inputrow, 1).Value ' ticker becomes these rows
            closeprice = ws.Cells(inputrow, 6).Value
            
            ' Debug.Print Ticker, closeprice
            ' Calculations
            opentoclose = closeprice - openprice
            volume = volume + ws.Cells(inputrow, 7).Value
            If openprice = 0 Then
            percentchange = 0 ' fixing divide by zero error
            
            Else
            
            percentchange = (opentoclose / openprice) ' calculate percentage
           
            End If
            
            ' Outputs
            ws.Cells(outputrow, 8).Value = Ticker ' row outputrow column H ' these rows become ticker
            ws.Cells(outputrow, 9).Value = opentoclose ' column I
            ws.Cells(outputrow, 10).Value = percentchange ' percent change goes to J
            ws.Cells(outputrow, 11).Value = volume

            ' Debug.Print Ticker, percentchange
            ' Continue
            openprice = ws.Cells(inputrow + 1, 3).Value 'opening price for NEXT stock
            outputrow = outputrow + 1 ' move down one row on output
            volume = 0 ' reset volume
            Else
            volume = volume + ws.Cells(inputrow, 7).Value ' summing up volume with each stock
                       
           
        End If
    ' coloring cells
    If ws.Cells(inputrow, 10).Value > 0 Then
    ws.Cells(inputrow, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(inputrow, 10).Value < 0 Then
    ws.Cells(inputrow, 10).Interior.ColorIndex = 3
    ElseIf ws.Cells(inputrow, 10).Value = 0 Then
    ws.Cells(inputrow, 10).Interior.ColorIndex = 2
    End If
    ' formatting to percentage
    Cells(inputrow, 10).NumberFormat = "0.00%"
    
    Range("H1") = "Ticker"
    Range("I1") = "Yearly Change"
    Range("J1") = "Percent Change"
    Range("K1") = "Total Volume"
    Next inputrow
Next ws
   


End Sub

