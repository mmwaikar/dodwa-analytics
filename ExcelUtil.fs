module ExcelUtil

open Types

let getPacktLibData =
    let sheet = SheetPacktLib()
    let data = sheet.Data
    let rows = List.ofSeq (data)

    let row2015 = rows.Item(1)
    let row2016 = rows.Item(2)
    let row2017 = rows.Item(3)
    let row2018 = rows.Item(4)
    let row2019 = rows.Item(5)

    let sales2015 =
        { Q1 = row2015.Q1 :?> double
          Q2 = row2015.Q2 :?> double
          Q3 = row2015.Q3 :?> double
          Q4 = row2015.Q4 :?> double }

    let sales2016 =
        { Q1 = row2016.Q1 :?> double
          Q2 = row2016.Q2 :?> double
          Q3 = row2016.Q3 :?> double
          Q4 = row2016.Q4 :?> double }

    let sales2017 =
        { Q1 = row2017.Q1 :?> double
          Q2 = row2017.Q2 :?> double
          Q3 = row2017.Q3 :?> double
          Q4 = row2017.Q4 :?> double }

    let sales2018 =
        { Q1 = row2018.Q1 :?> double
          Q2 = row2018.Q2 :?> double
          Q3 = row2018.Q3 :?> double
          Q4 = row2018.Q4 :?> double }

    let sales2019 =
        { Q1 = row2019.Q1 :?> double
          Q2 = row2019.Q2 :?> double
          Q3 = row2019.Q3 :?> double
          Q4 = row2019.Q4 :?> double }

    { QuarterWiseYearlySales.Y2015 = sales2015
      QuarterWiseYearlySales.Y2016 = sales2016
      QuarterWiseYearlySales.Y2017 = sales2017
      QuarterWiseYearlySales.Y2018 = sales2018
      QuarterWiseYearlySales.Y2019 = sales2019 }

let getOthersData =
    let sheet = SheetOthers()
    let data = sheet.Data
    let rows = List.ofSeq (data)

    let row2015 = rows.Item(1)
    let row2016 = rows.Item(2)
    let row2017 = rows.Item(3)
    let row2018 = rows.Item(4)
    let row2019 = rows.Item(5)

    let sales2015 =
        { Q1 = row2015.Q1 :?> double
          Q2 = row2015.Q2 :?> double
          Q3 = row2015.Q3 :?> double
          Q4 = row2015.Q4 :?> double }

    let sales2016 =
        { Q1 = row2016.Q1 :?> double
          Q2 = row2016.Q2 :?> double
          Q3 = row2016.Q3 :?> double
          Q4 = row2016.Q4 :?> double }

    let sales2017 =
        { Q1 = row2017.Q1 :?> double
          Q2 = row2017.Q2 :?> double
          Q3 = row2017.Q3 :?> double
          Q4 = row2017.Q4 :?> double }

    let sales2018 =
        { Q1 = row2018.Q1 :?> double
          Q2 = row2018.Q2 :?> double
          Q3 = row2018.Q3 :?> double
          Q4 = row2018.Q4 :?> double }

    let sales2019 =
        { Q1 = row2019.Q1 :?> double
          Q2 = row2019.Q2 :?> double
          Q3 = row2019.Q3 :?> double
          Q4 = row2019.Q4 :?> double }

    { QuarterWiseYearlySales.Y2015 = sales2015
      QuarterWiseYearlySales.Y2016 = sales2016
      QuarterWiseYearlySales.Y2017 = sales2017
      QuarterWiseYearlySales.Y2018 = sales2018
      QuarterWiseYearlySales.Y2019 = sales2019 }

let getPBooksData =
    let sheet = SheetPbooks()
    let data = sheet.Data
    let rows = List.ofSeq (data)

    let row2015 = rows.Item(1)
    let row2016 = rows.Item(2)
    let row2017 = rows.Item(3)
    let row2018 = rows.Item(4)
    let row2019 = rows.Item(5)

    let sales2015 =
        { Q1 = row2015.Q1 :?> double
          Q2 = row2015.Q2 :?> double
          Q3 = row2015.Q3 :?> double
          Q4 = row2015.Q4 :?> double }

    let sales2016 =
        { Q1 = row2016.Q1 :?> double
          Q2 = row2016.Q2 :?> double
          Q3 = row2016.Q3 :?> double
          Q4 = row2016.Q4 :?> double }

    let sales2017 =
        { Q1 = row2017.Q1 :?> double
          Q2 = row2017.Q2 :?> double
          Q3 = row2017.Q3 :?> double
          Q4 = row2017.Q4 :?> double }

    let sales2018 =
        { Q1 = row2018.Q1 :?> double
          Q2 = row2018.Q2 :?> double
          Q3 = row2018.Q3 :?> double
          Q4 = row2018.Q4 :?> double }

    let sales2019 =
        { Q1 = row2019.Q1 :?> double
          Q2 = row2019.Q2 :?> double
          Q3 = row2019.Q3 :?> double
          Q4 = row2019.Q4 :?> double }

    { QuarterWiseYearlySales.Y2015 = sales2015
      QuarterWiseYearlySales.Y2016 = sales2016
      QuarterWiseYearlySales.Y2017 = sales2017
      QuarterWiseYearlySales.Y2018 = sales2018
      QuarterWiseYearlySales.Y2019 = sales2019 }

let getEBooksData =
    let sheet = SheetEbooks()
    let data = sheet.Data
    let rows = List.ofSeq (data)

    let row2015 = rows.Item(1)
    let row2016 = rows.Item(2)
    let row2017 = rows.Item(3)
    let row2018 = rows.Item(4)
    let row2019 = rows.Item(5)

    let sales2015 =
        { Q1 = row2015.Q1 :?> double
          Q2 = row2015.Q2 :?> double
          Q3 = row2015.Q3 :?> double
          Q4 = row2015.Q4 :?> double }

    let sales2016 =
        { Q1 = row2016.Q1 :?> double
          Q2 = row2016.Q2 :?> double
          Q3 = row2016.Q3 :?> double
          Q4 = row2016.Q4 :?> double }

    let sales2017 =
        { Q1 = row2017.Q1 :?> double
          Q2 = row2017.Q2 :?> double
          Q3 = row2017.Q3 :?> double
          Q4 = row2017.Q4 :?> double }

    let sales2018 =
        { Q1 = row2018.Q1 :?> double
          Q2 = row2018.Q2 :?> double
          Q3 = row2018.Q3 :?> double
          Q4 = row2018.Q4 :?> double }

    let sales2019 =
        { Q1 = row2019.Q1 :?> double
          Q2 = row2019.Q2 :?> double
          Q3 = row2019.Q3 :?> double
          Q4 = row2019.Q4 :?> double }

    { QuarterWiseYearlySales.Y2015 = sales2015
      QuarterWiseYearlySales.Y2016 = sales2016
      QuarterWiseYearlySales.Y2017 = sales2017
      QuarterWiseYearlySales.Y2018 = sales2018
      QuarterWiseYearlySales.Y2019 = sales2019 }

let getYearlySales (quarterWiseYearlySales: QuarterWiseYearlySales) =
    let totalSales2015 = quarterWiseYearlySales.Y2015.Q1 + quarterWiseYearlySales.Y2015.Q2 + quarterWiseYearlySales.Y2015.Q3 + quarterWiseYearlySales.Y2015.Q4
    let totalSales2016 = quarterWiseYearlySales.Y2016.Q1 + quarterWiseYearlySales.Y2016.Q2 + quarterWiseYearlySales.Y2016.Q3 + quarterWiseYearlySales.Y2016.Q4
    let totalSales2017 = quarterWiseYearlySales.Y2017.Q1 + quarterWiseYearlySales.Y2017.Q2 + quarterWiseYearlySales.Y2017.Q3 + quarterWiseYearlySales.Y2017.Q4
    let totalSales2018 = quarterWiseYearlySales.Y2018.Q1 + quarterWiseYearlySales.Y2018.Q2 + quarterWiseYearlySales.Y2018.Q3 + quarterWiseYearlySales.Y2018.Q4
    let totalSales2019 = quarterWiseYearlySales.Y2019.Q1 + quarterWiseYearlySales.Y2019.Q2 + quarterWiseYearlySales.Y2019.Q3 + quarterWiseYearlySales.Y2019.Q4

    { YearlySales.Y2015 = totalSales2015
      YearlySales.Y2016 = totalSales2016
      YearlySales.Y2017 = totalSales2017
      YearlySales.Y2018 = totalSales2018
      YearlySales.Y2019 = totalSales2019 }

let getTotalSales (yearlySales: YearlySales) =
    yearlySales.Y2015 + yearlySales.Y2016 + yearlySales.Y2017 + yearlySales.Y2018 + yearlySales.Y2019

let getYearlySalesForAll (packtQuarterlySales: QuarterWiseYearlySales, othersQuarterlySales: QuarterWiseYearlySales,
                          pBooksQuarterlySales: QuarterWiseYearlySales, eBooksQuarterlySales: QuarterWiseYearlySales) =
    let packtYearlySales = getYearlySales packtQuarterlySales
    let othersYearlySales = getYearlySales othersQuarterlySales
    let pBooksYearlySales = getYearlySales pBooksQuarterlySales
    let eBooksYearlySales = getYearlySales eBooksQuarterlySales
    (packtYearlySales, othersYearlySales, pBooksYearlySales, eBooksYearlySales)

let getYearlySalesLabelledByYear (yearlySales: YearlySales) =
    [ "2015", yearlySales.Y2015
      "2016", yearlySales.Y2016
      "2017", yearlySales.Y2017
      "2018", yearlySales.Y2018
      "2019", yearlySales.Y2019 ]

let getYearlySalesLabelledByYearForAll (packtYearlySales: YearlySales, othersYearlySales: YearlySales,
                                        pBooksYearlySales: YearlySales, eBooksYearlySales: YearlySales) =
    let packtLibSalesWithLabel = getYearlySalesLabelledByYear packtYearlySales
    let othersSalesWithLabel = getYearlySalesLabelledByYear othersYearlySales
    let pBooksSalesWithLabel = getYearlySalesLabelledByYear pBooksYearlySales
    let eBooksSalesWithLabel = getYearlySalesLabelledByYear eBooksYearlySales
    (packtLibSalesWithLabel, othersSalesWithLabel, pBooksSalesWithLabel, eBooksSalesWithLabel)

let getYearlySalesLabelledByItemForAll (packtYearlySales: YearlySales, othersYearlySales: YearlySales,
                                          pBooksYearlySales: YearlySales, eBooksYearlySales: YearlySales) =
    let salesWith2015Label = [ "2015", packtYearlySales.Y2015
                               "2015", othersYearlySales.Y2015
                               "2015", pBooksYearlySales.Y2015
                               "2015", eBooksYearlySales.Y2015 ]

    let salesWith2016Label = [ "2016", packtYearlySales.Y2016 
                               "2016", othersYearlySales.Y2016
                               "2016", pBooksYearlySales.Y2016 
                               "2016", eBooksYearlySales.Y2016 ]

    let salesWith2017Label = [ "2017", packtYearlySales.Y2017
                               "2017", othersYearlySales.Y2017
                               "2017", pBooksYearlySales.Y2017
                               "2017", eBooksYearlySales.Y2017 ]

    let salesWith2018Label = [ "2018", packtYearlySales.Y2018
                               "2018", othersYearlySales.Y2018
                               "2018", pBooksYearlySales.Y2018
                               "2018", eBooksYearlySales.Y2018 ]

    let salesWith2019Label = [ "2019", packtYearlySales.Y2019
                               "2019", othersYearlySales.Y2019
                               "2019", pBooksYearlySales.Y2019
                               "2019", eBooksYearlySales.Y2019 ]
                               
    (salesWith2015Label, salesWith2016Label, salesWith2017Label, salesWith2018Label, salesWith2019Label)
