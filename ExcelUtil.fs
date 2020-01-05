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

    { Y2015 = sales2015
      Y2016 = sales2016
      Y2017 = sales2017
      Y2018 = sales2018
      Y2019 = sales2019 }

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

    { Y2015 = sales2015
      Y2016 = sales2016
      Y2017 = sales2017
      Y2018 = sales2018
      Y2019 = sales2019 }

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

    { Y2015 = sales2015
      Y2016 = sales2016
      Y2017 = sales2017
      Y2018 = sales2018
      Y2019 = sales2019 }

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

    { Y2015 = sales2015
      Y2016 = sales2016
      Y2017 = sales2017
      Y2018 = sales2018
      Y2019 = sales2019 }            