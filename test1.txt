Sub StokDashboard()
    Dim StokTbl As ListObject
    Dim StokData As Range
    Dim StokChart As Chart
    Dim StokChartRange As Range
    Dim StokChartTitle As String
    
    ' Stok tablosu oluşturma
    Set StokData = Sheets("Sheet1").Range("A1:C101")
    Set StokTbl = Sheets("Sheet1").ListObjects.Add(xlSrcRange, StokData, , xlYes)
    StokTbl.TableStyle = "TableStyleMedium2"
    StokTbl.Range.Columns(1).ColumnWidth = 14
    StokTbl.Range.Columns(2).ColumnWidth = 30
    
    ' Stok grafiği oluşturma
    Set StokChart = Sheets("Sheet1").Shapes.AddChart2(251, xlColumnClustered).Chart
    StokChart.Parent.Name = "Stok Grafik"
    StokChart.SetSourceData StokTbl.Range.Columns(3)
    StokChart.HasLegend = False
    StokChartTitle = "Stok Durumu"
    StokChart.ChartTitle.Text = StokChartTitle
    StokChart.Axes(xlCategory).TickLabels.Font.Size = 10
    Set StokChartRange = StokTbl.Range.Columns(1).Resize(StokTbl.ListRows.Count + 1, 1)
    StokChart.Axes(xlCategory).CategoryNames = StokChartRange
    StokChart.Axes(xlValue).TickLabels.NumberFormat = "#,##0.00"
    StokChart.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    StokChart.ChartArea.Format.Line.Visible = False
    StokChart.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    StokChart.PlotArea.Format.Line.Visible = False
    StokChart.ChartArea.Border.LineStyle = xlNone
    StokChart.PlotArea.Border.LineStyle = xlNone
    
    ' Stok takip tablosu oluşturma
    Dim StokTakipTbl As ListObject
    Dim StokTakipData As Range
    Dim StokTakipChart As Chart
    Dim StokTakipChartRange As Range
    Dim StokTakipChartTitle As String
    
    Set StokTakipData = Sheets("Sheet1").Range("E1:H101")
    Set StokTakipTbl = Sheets("Sheet1").ListObjects.Add(xlSrcRange, StokTakipData, , xlYes)
    StokTakipTbl.TableStyle = "TableStyleMedium2"
    StokTakipTbl.Range.Columns(1).ColumnWidth = 14
    StokTakipTbl.Range.Columns(2).ColumnWidth = 30
    StokTakipTbl.Range.Columns(3).ColumnWidth = 14
    StokTakipTbl.Range.Columns(4).ColumnWidth = 14
    
    ' Stok takip grafiği oluşturma
    Set StokTakipChart = Sheets("Sheet1").Shapes.AddChart2(251, xlLineMarkers).Chart
    StokTakipChart.Parent.Name = "Stok Takip Grafik"
    StokTakipChart.SetSourceData StokTak
