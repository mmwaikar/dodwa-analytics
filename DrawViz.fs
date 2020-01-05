module DrawViz

open Types
open ExcelUtil
open XPlot.GoogleCharts

let quartersInAYear = [ "Q1"; "Q2"; "Q3"; "Q4" ]
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
            (title = titlePrefix, hAxis = Axis(title = "Year", titleTextStyle = TextStyle(color = "red")),
             vAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "green")), seriesType = "bars")

    [ q1Sales; q2Sales; q3Sales; q4Sales ]
    |> Chart.Combo
    |> Chart.WithOptions options
    |> Chart.WithLabels quartersInAYear

let drawYearlyChart (packtQuarterlySales: QuarterWiseYearlySales, othersQuarterlySales: QuarterWiseYearlySales,
                     pBooksQuarterlySales: QuarterWiseYearlySales, eBooksQuarterlySales: QuarterWiseYearlySales) =
    let packtYearlySales = getYearlySales packtQuarterlySales
    let othersYearlySales = getYearlySales othersQuarterlySales
    let pBooksYearlySales = getYearlySales pBooksQuarterlySales
    let eBooksYearlySales = getYearlySales eBooksQuarterlySales

    let packtLibSales =
        [ "2015", packtYearlySales.Y2015
          "2016", packtYearlySales.Y2016
          "2017", packtYearlySales.Y2017
          "2018", packtYearlySales.Y2018
          "2019", packtYearlySales.Y2019 ]

    let othersSales =
        [ "2015", othersYearlySales.Y2015
          "2016", othersYearlySales.Y2016
          "2017", othersYearlySales.Y2017
          "2018", othersYearlySales.Y2018
          "2019", othersYearlySales.Y2019 ]

    let pBooksSales =
        [ "2015", pBooksYearlySales.Y2015
          "2016", pBooksYearlySales.Y2016
          "2017", pBooksYearlySales.Y2017
          "2018", pBooksYearlySales.Y2018
          "2019", pBooksYearlySales.Y2019 ]

    let eBooksSales =
        [ "2015", eBooksYearlySales.Y2015
          "2016", eBooksYearlySales.Y2016
          "2017", eBooksYearlySales.Y2017
          "2018", eBooksYearlySales.Y2018
          "2019", eBooksYearlySales.Y2019 ]

    let options =
        Options
            (title = "Sales of Data Oriented Development with AngularJS",
             hAxis = Axis(title = "Year", titleTextStyle = TextStyle(color = "red")),
             vAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "green")), seriesType = "bars")

    [ packtLibSales; othersSales; pBooksSales; eBooksSales ]
    |> Chart.Combo
    |> Chart.WithOptions options
    |> Chart.WithLabels items

let drawTotalChart (packtQuarterlySales: QuarterWiseYearlySales, othersQuarterlySales: QuarterWiseYearlySales,
                     pBooksQuarterlySales: QuarterWiseYearlySales, eBooksQuarterlySales: QuarterWiseYearlySales) =
    let packtYearlySales = getYearlySales packtQuarterlySales
    let othersYearlySales = getYearlySales othersQuarterlySales
    let pBooksYearlySales = getYearlySales pBooksQuarterlySales
    let eBooksYearlySales = getYearlySales eBooksQuarterlySales

    let totalPacktSales = [ items.Item(0), (getTotalSales packtYearlySales) ]
    let totalOthersSales = [ items.Item(1), (getTotalSales othersYearlySales) ]
    let totalPBookSales = [ items.Item(2), (getTotalSales pBooksYearlySales) ]
    let totalEBookSales = [ items.Item(3), (getTotalSales eBooksYearlySales) ]

    let options =
        Options
            (title = "Sales of Data Oriented Development with AngularJS",
             hAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "green")),
             vAxis = Axis(title = "Item", titleTextStyle = TextStyle(color = "red")))

    [ totalPacktSales; totalOthersSales; totalPBookSales; totalEBookSales ]
    |> Chart.Bar
    |> Chart.WithOptions options
    |> Chart.WithLabels items
