module DrawViz

open Types
open XPlot.GoogleCharts

let drawChart (sales: YearlySales, titlePrefix: string) =
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
            (title = titlePrefix + " Sales", hAxis = Axis(title = "Year", titleTextStyle = TextStyle(color = "red")),
             vAxis = Axis(title = "Sales", titleTextStyle = TextStyle(color = "blue")), seriesType = "bars")

    [ q1Sales; q2Sales; q3Sales; q4Sales ]
    |> Chart.Combo
    |> Chart.WithOptions options
    |> Chart.WithLabels [ "Q1"; "Q2"; "Q3"; "Q4" ]
