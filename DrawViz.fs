module DrawViz

open Types
open ExcelUtil
open XPlot.GoogleCharts

let quartersInAYear = [ "Q1"; "Q2"; "Q3"; "Q4" ]
let yearsInSales = [ "2015"; "2016"; "2017"; "2018"; "2019" ]
let items = [ "Subscription PacktLib"; "Subscription Others"; "Print Books"; "Ebooks" ]

let drawQuarterWiseChart (sales: QuarterWiseYearlySales, titlePrefix: string) =
    let q1Sales =
        [ "2015", sales.Y2015.Q1
          "2016", sales.Y2016.Q1
          "2017", sales.Y2017.Q1
          "2018", sales.Y2018.Q1
          "2019", sales.Y2019.Q1 ]

    let q2Sales =
        [ "2015", sales.Y2015.Q2
          "2016", sales.Y2016.Q2
          "2017", sales.Y2017.Q2
          "2018", sales.Y2018.Q2
          "2019", sales.Y2019.Q2 ]

    let q3Sales =
        [ "2015", sales.Y2015.Q3
          "2016", sales.Y2016.Q3
          "2017", sales.Y2017.Q3
          "2018", sales.Y2018.Q3
          "2019", sales.Y2019.Q3 ]

    let q4Sales =
        [ "2015", sales.Y2015.Q4
          "2016", sales.Y2016.Q4
          "2017", sales.Y2017.Q4
          "2018", sales.Y2018.Q4
          "2019", sales.Y2019.Q4 ]

    let options =
        Options
            (title = titlePrefix + " (per quarter per year)", hAxis = Axis(title = "Year", titleTextStyle = TextStyle(color = "red")),
             vAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "green")), seriesType = "bars")

    [ q1Sales; q2Sales; q3Sales; q4Sales ]
    |> Chart.Combo
    |> Chart.WithOptions options
    |> Chart.WithLabels quartersInAYear

let drawYearlyChart (packtQuarterlySales: QuarterWiseYearlySales, othersQuarterlySales: QuarterWiseYearlySales,
                     pBooksQuarterlySales: QuarterWiseYearlySales, eBooksQuarterlySales: QuarterWiseYearlySales) =
    let (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales) =
        getYearlySalesForAll (packtQuarterlySales, othersQuarterlySales, pBooksQuarterlySales, eBooksQuarterlySales)

    let (packtLibSalesWithLabel, othersSalesWithLabel, pBooksSalesWithLabel, eBooksSalesWithLabel) =
        getYearlySalesLabelledByYearForAll (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales)

    let options =
        Options
            (title = "Sales of Data Oriented Development with AngularJS",
             hAxis = Axis(title = "Year", titleTextStyle = TextStyle(color = "red")),
             vAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "green")), seriesType = "bars")

    [ packtLibSalesWithLabel; othersSalesWithLabel; pBooksSalesWithLabel; eBooksSalesWithLabel ]
    |> Chart.Combo
    |> Chart.WithOptions options
    |> Chart.WithLabels items

let drawLineChart (packtQuarterlySales: QuarterWiseYearlySales, othersQuarterlySales: QuarterWiseYearlySales,
                     pBooksQuarterlySales: QuarterWiseYearlySales, eBooksQuarterlySales: QuarterWiseYearlySales) =
    let (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales) =
        getYearlySalesForAll (packtQuarterlySales, othersQuarterlySales, pBooksQuarterlySales, eBooksQuarterlySales)

    let (packtLibSalesWithLabel, othersSalesWithLabel, pBooksSalesWithLabel, eBooksSalesWithLabel) =
        getYearlySalesLabelledByYearForAll (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales)

    let options =
        Options
            (title = "Sales trend of Data Oriented Development with AngularJS",
             legend = Legend(position = "bottom"), height = 640, width = 1280, pointShape = "triangle")

    [ packtLibSalesWithLabel; othersSalesWithLabel; pBooksSalesWithLabel; eBooksSalesWithLabel ]
    |> Chart.Line
    |> Chart.WithOptions options
    |> Chart.WithLabels items

let drawTotalChart (packtQuarterlySales: QuarterWiseYearlySales, othersQuarterlySales: QuarterWiseYearlySales,
                    pBooksQuarterlySales: QuarterWiseYearlySales, eBooksQuarterlySales: QuarterWiseYearlySales) =
    let (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales) =
        getYearlySalesForAll (packtQuarterlySales, othersQuarterlySales, pBooksQuarterlySales, eBooksQuarterlySales)

    let totalPacktSales = [ items.Item(0), (getTotalSales packtYearlySales) ]
    let totalOthersSales = [ items.Item(1), (getTotalSales othersYearlySales) ]
    let totalPBookSales = [ items.Item(2), (getTotalSales pBooksYearlySales) ]
    let totalEBookSales = [ items.Item(3), (getTotalSales eBooksYearlySales) ]

    let options =
        Options
            (title = "Total sales of Data Oriented Development with AngularJS",
             hAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "green")),
             vAxis = Axis(title = "Item", titleTextStyle = TextStyle(color = "red")))

    [ totalPacktSales; totalOthersSales; totalPBookSales; totalEBookSales ]
    |> Chart.Bar
    |> Chart.WithOptions options
    |> Chart.WithLabels items

let drawYearWiseStackedChart (packtQuarterlySales: QuarterWiseYearlySales, othersQuarterlySales: QuarterWiseYearlySales,
                              pBooksQuarterlySales: QuarterWiseYearlySales, eBooksQuarterlySales: QuarterWiseYearlySales) =
    let (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales) =
        getYearlySalesForAll (packtQuarterlySales, othersQuarterlySales, pBooksQuarterlySales, eBooksQuarterlySales)

    let options =
        Options
            (title = "Sales of Data Oriented Development with AngularJS (stacked)", isStacked = true,
             hAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "green")),
             vAxis = Axis(title = "Year", titleTextStyle = TextStyle(color = "red")))

    let (packtLibSalesWithLabel, othersSalesWithLabel, pBooksSalesWithLabel, eBooksSalesWithLabel) =
        getYearlySalesLabelledByYearForAll (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales)

    [ packtLibSalesWithLabel; othersSalesWithLabel; pBooksSalesWithLabel; eBooksSalesWithLabel ]
    |> Chart.Bar
    |> Chart.WithOptions options
    |> Chart.WithLabels items

/// does not work yet
let drawItemWiseStackedChart (packtQuarterlySales: QuarterWiseYearlySales, othersQuarterlySales: QuarterWiseYearlySales,
                              pBooksQuarterlySales: QuarterWiseYearlySales, eBooksQuarterlySales: QuarterWiseYearlySales) =
    let (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales) =
        getYearlySalesForAll (packtQuarterlySales, othersQuarterlySales, pBooksQuarterlySales, eBooksQuarterlySales)

    let (salesWith2015Label, salesWith2016Label, salesWith2017Label, salesWith2018Label, salesWith2019Label) =
        getYearlySalesLabelledByItemForAll (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales)

    let options =
        Options
            (title = "Sales of Data Oriented Development with AngularJS", isStacked = true,
             hAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "green")),
             vAxis = Axis(title = "Item", titleTextStyle = TextStyle(color = "red")))

    [ salesWith2015Label; salesWith2016Label; salesWith2017Label; salesWith2018Label; salesWith2019Label ]
    |> Chart.Bar
    |> Chart.WithOptions options
    |> Chart.WithLabels yearsInSales