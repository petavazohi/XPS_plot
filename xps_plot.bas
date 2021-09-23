Attribute VB_Name = "Module1"
Sub plot_xps()
Attribute plot_xps.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' plot_xps Macro
'

'
    Dim lCol As String
    Dim lRow As Integer
    Dim Emin As Long
    Dim Emax As Long
    Dim Cmin As Long
    Dim Cmax As Long
    Dim BE As Range
    Dim Counts As Range
    Dim name As String
    Dim cht As Shape
    Dim height As Integer
    Dim width As Integer
    Dim Color As Long
    Dim white As Long
    Dim i As Integer
    Dim axx As Axis
    Dim axy As Axis
    
    ' size of the chart
    width = 500
    height = 300
    
    ' name of the sheet to be used later for the chart as well
    name = ActiveSheet.name
    
    ' This for loop finds column with the name Envelope
    For Each c In Range("A4:Z4")
        c.Value = Split(c.Value, "_")
        If InStr(c.Value, "Envelope") > 0 Then
           lCol = Trim(Replace(c.Address, "$", ""))
           lCol = Left(lCol, 1)
        End If
    Next c
    
    ' This line finds the last row used in this sheet
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
    ' Maximum and Minimum energies present in the plot
    Set BE = Range("B5:B" & lRow)
    Emin = Application.WorksheetFunction.Min(BE)
    Emax = Application.WorksheetFunction.Max(BE)
    
    ' Maximum and Minimum CPS present in the plot
    Set Counts = Range("C5:" & lCol & lRow)
    Cmin = Application.WorksheetFunction.Min(Counts)
    Cmax = Application.WorksheetFunction.Max(Counts)
    
    ' Ploting style 240
    ActiveSheet.Range("B4", lCol & lRow).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    ActiveChart.SetSourceData Source:=Range(name & "!$B$4:$" & lCol & "$" & lRow)
    
    ' Setting x limits
    ActiveChart.Axes(xlCategory).MinimumScale = Emin
    ActiveChart.Axes(xlCategory).MaximumScale = Emax
    
    ' Setting y limits to 95% minimum of CPS and 120% maximum CPS
    ActiveChart.Axes(xlValue).MinimumScale = Application.WorksheetFunction.RoundDown(Cmin * 0.95, 0)
    ActiveChart.Axes(xlValue).MaximumScale = Application.WorksheetFunction.RoundUp(Cmax * 1.2, 0)
    
    ' Changing the title name to the sheet name
    ActiveChart.ChartTitle.Text = name
    
    ' Changing the chart name to the sheet name
    Set cht = ActiveSheet.Shapes(1)
    cht.name = name
    
    ' Adding Major and minor grid lines to the plot
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
    
    ' Changing the font to Times New Roman
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "Times New Roman"
        .name = "Times New Roman"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
    End With
    

    
    ' Changing the size of the plot
    With ActiveSheet.ChartObjects(name)
        .height = height ' resize
        .width = width  ' resize
        .Top = 20    ' reposition
        .Left = 50   ' reposition
    End With
    
    ' Formating major and minor ticks to be cross and inside, respectivley
    ActiveChart.Axes(xlCategory).MajorTickMark = xlCross
    ActiveChart.Axes(xlCategory).MinorTickMark = xlInside
    ActiveChart.Axes(xlValue).MajorTickMark = xlCross
    ActiveChart.Axes(xlValue).MinorTickMark = xlInside
    
    ' Legend Settings
    ActiveChart.Legend.Select
    Selection.Position = xlCorner
    'ActiveChart.Legend.Left = 0.04 * width 'width / 10
    'ActiveChart.Legend.Top = 0 'height / 10
    'ActiveChart.Legend.Select
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Size = 10.5
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Weight = 1.25
    End With
    
    ' Changing the plot area to be almost the same as the plot size
    ActiveChart.PlotArea.Left = 0.04 * width '32.655
    ActiveChart.PlotArea.Top = 0 '4.982
    ActiveChart.PlotArea.height = height * 0.94 '401.105
    ActiveChart.PlotArea.width = width  '710.344
    
    ' Chaning the color of x and y axis
    ActiveChart.Axes(xlCategory).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    
    ActiveChart.Axes(xlValue).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    

    ' Adding y label
    ActiveSheet.ChartObjects(name).Activate
    With ActiveChart.Axes(xlValue)
        .HasTitle = True
        With .AxisTitle
            .Caption = "Counts per Second"
            .Font.name = "Times New Roman"
            .Font.Size = 12
        End With
    End With
    
    ' Adding x label
    ActiveSheet.ChartObjects(name).Activate
    With ActiveChart.Axes(xlCategory)
        .HasTitle = True
        With .AxisTitle
            .Caption = "Binding Energy (eV)"
            .Font.name = "Times New Roman"
            .Font.Size = 12
        End With
    End With
    
    ' cheking white color
    For Each c In Range("AZ100:AZ101")
        white = c.DisplayFormat.Interior.Color
        
    Next c
    
    ' Changing the color of the lines to be the same as the color of column header
    For Each c In Range("D4:" & lCol & "4")
        Color = c.DisplayFormat.Interior.Color
        ActiveChart.FullSeriesCollection(2 + i).Select
        With Selection.Format.Line
            .Visible = msoTrue
            If Color <> white Then
                .ForeColor.RGB = Color
            End If
            .Transparency = 0
            .Weight = 2
        End With
        i = i + 1
    Next c
    
    ' Changing the color of the original data to black and dashed line
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .DashStyle = msoLineSysDash
        .ForeColor.RGB = RGB(0, 0, 0)
        .Visible = msoTrue
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    


    ' Changing the font size of y axis to 12
    ActiveSheet.ChartObjects(name).Activate
    Set axy = ActiveChart.Axes(xlValue)
    axy.TickLabels.Font.Size = 12
    
    ' Changing the font size of x axis to 12
    ActiveSheet.ChartObjects(name).Activate
    Set axx = ActiveChart.Axes(xlCategory)
    axx.TickLabels.Font.Size = 12
    
    
    'Selection.Format.TextFrame2.TextRange.Font.Size = 12
    'ActiveChart.Axes(xlValue).Select
    'Selection.Format.TextFrame2.TextRange.Font.Size = 12
    
    ' Changing the font size of x axis to 12
    'ActiveSheet.ChartObjects(name).Activate
    'ActiveChart.Axes(xlCategory).Select
    'Selection.Format.TextFrame2.TextRange.Font.Size = 12

End Sub
