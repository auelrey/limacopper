limacopper
==========
Private Sub CommandButton1_Click()
If ComboBox1.Value = "Verizon wireless" And ComboBox2.Value = "AT&T" Then

    

newdata = Sheets("Data").Range("A1:C7").Value
Sheets("Sheet1").Range("O7:Q13").Value = newdata
Unload Me


Range("A1:C7").Select
    Sheets("Sheet1").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=Range("Data!$A$1:$C$7")

ElseIf ComboBox1.Value = "PTR" And ComboBox2.Value = "SHI" Then



    Dim iChart As Long 'chart counter
    Dim nCharts As Long 'total charts

    nCharts = ActiveSheet.ChartObjects.Count
    
    'now we are going to loop through all the charts
    For iChart = 1 To nCharts
        With Sheets(1).ChartObjects(1)
            .Delete 'and remove them 1 by 1
        End With
    Next
    
    With ActiveSheet
        .Shapes.AddChart.Select
    End With
    
    With ActiveChart
        Do Until .SeriesCollection.Count = 0
            .SeriesCollection(1).Delete
        Loop
        
        .ChartType = xlColumnClustered
        .ChartArea.Select
        .HasTitle = True
        .ChartTitle.Text = ""
    End With
    
    'now we choose where to use our data points to create our chart
    With ActiveChart
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = Sheets("AnalyseData").Cells(1, 2).Value 'legend
        .SeriesCollection(1).Values = Sheets("AnalyseData").Range("B2:B10") 'values
        .SeriesCollection(1).XValues = Sheets("AnalyseData").Range("A2:A10") 'x axis
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = Sheets("AnalyseData").Cells(1, 3).Value 'legend
        .SeriesCollection(2).Values = Sheets("AnalyseData").Range("C2:C10") 'values
        .SeriesCollection(2).XValues = Sheets("AnalyseData").Range("A2:A10") 'x axis
    End With
    
    
    With ActiveChart
        .Shapes.AddChart.Select
    End With
    
    With ActiveChart
        Do Until .SeriesCollection.Count = 0
            .SeriesCollection(1).Delete
        Loop
        
        .ChartType = xlColumnClustered
        .ChartArea.Select
        .HasTitle = True
        .ChartTitle.Text = ""
    End With
    
    'now we choose where to use our data points to create our chart
    With ActiveChart
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = Sheets("AnalyseData").Cells(1, 2).Value 'legend
        .SeriesCollection(1).Values = Sheets("AnalyseData").Range("F2:F5") 'values
        .SeriesCollection(1).XValues = Sheets("AnalyseData").Range("E2:E5") 'x axis
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = Sheets("AnalyseData").Cells(1, 3).Value 'legend
        .SeriesCollection(2).Values = Sheets("AnalyseData").Range("G2:G5") 'values
        .SeriesCollection(2).XValues = Sheets("AnalyseData").Range("E2:E5") 'x axis
    End With
    
    
    With ActiveChart
        .Shapes.AddChart.Select
    End With
    
    With ActiveChart
        Do Until .SeriesCollection.Count = 0
            .SeriesCollection(1).Delete
        Loop
        
        .ChartType = xlColumnClustered
        .ChartArea.Select
        .HasTitle = True
        .ChartTitle.Text = ""
    End With
    
    'now we choose where to use our data points to create our chart
    With ActiveChart
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = Sheets("AnalyseData").Cells(1, 2).Value 'legend
        .SeriesCollection(1).Values = Sheets("AnalyseData").Range("K2:K3") 'values
        .SeriesCollection(1).XValues = Sheets("AnalyseData").Range("J2:J3") 'x axis
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = Sheets("AnalyseData").Cells(1, 3).Value 'legend
        .SeriesCollection(2).Values = Sheets("AnalyseData").Range("L2:L3") 'values
        .SeriesCollection(2).XValues = Sheets("AnalyseData").Range("J2:J3") 'x axis
    End With

    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double
    Dim dWidth As Double
    Dim nColumns As Long
    
    dTop = 15
    dLeft = 300
    dHeight = 300
    dWidth = 450
    nColumns = 1
    nCharts = Sheets("Analyse").ChartObjects.Count
    
    For iChart = 1 To nCharts
        With Sheets(1).ChartObjects(iChart)
            .Height = dHeight
            .Width = dWidth
            .Left = dLeft + ((iChart - 1) Mod nColumns) * dWidth
            .Top = dTop + Int((iChart - 1) / nColumns) * dHeight
        End With
    Next
    
    selectedData = Sheets("Analyse").Range("O2:Q11").Value
    Sheets(1).Range("Q2:S11").Value = selectedData
    
    selectedData2 = Sheets("Analyse").Range("O22:Q26").Value
    Sheets(1).Range("Q22:S26").Value = selectedData2
    
    selectedData3 = Sheets("Analyse").Range("O42:Q44").Value
    Sheets(1).Range("Q42:S44").Value = selectedData3
    
    Sheets(1).Range("Q2:S11").Cells.Font.Color = RGB(255, 0, 0)
    Sheets(1).Range("Q2:S2").Interior.Color = RGB(202, 255, 112)
    Sheets(1).Range("Q3:S11").Interior.Color = RGB(162, 205, 90)
    
    Sheets(1).Range("Q22:S26").Cells.Font.Color = RGB(255, 0, 0)
    Sheets(1).Range("Q22:S22").Interior.Color = RGB(202, 255, 112)
    Sheets(1).Range("Q23:S26").Interior.Color = RGB(162, 205, 90)
    
    Sheets(1).Range("Q42:S44").Cells.Font.Color = RGB(255, 0, 0)
    Sheets(1).Range("Q42:S42").Interior.Color = RGB(202, 255, 112)
    Sheets(1).Range("Q43:S44").Interior.Color = RGB(162, 205, 90)
    Unload Me

End If

    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub
Private Sub UserForm_Initialize()
ComboBox1.List = Sheets("sheet2").Range("B1:B4").Value

ComboBox2.List = Sheets("sheet2").Range("B1:B4").Value

End Sub
