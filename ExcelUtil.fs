module ExcelUtil

open FSharp.Interop.Excel

let PacktLib = "Subscription PacktLib"
let Others = "Subscription Others"
let PrintBooks = "Book copies Print"
let EBooks = "Book copies Ebook"

type BookCopies =
    { Print: int32
      EBook: int32 }

type Subscription =
    { PacktLib: int32
      Others: int32 }

type Sales = 
    { Subscriptions: Subscription
      Books: BookCopies }

type Sheet2015 = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "2015">
type Sheet2016 = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "2016">
type Sheet2017 = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "2017">
type Sheet2018 = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "2018">
type Sheet2019 = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "2019">

let years = seq { 2015 .. 2019 }

// let getSheetType year =
//     ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", year.ToString>

// let readSheet year =
//     let rows = match year with
//                 | 2015 -> Sheet2015().Data
//                 | 2016 -> Sheet2016().Data
//                 | 2017 -> Sheet2017().Data
//                 | 2018 -> Sheet2018().Data
//                 | 2019 -> Sheet2019().Data
//                 | _ -> Sheet2020().Data

//     let packtLibRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = PacktLib)
//     let othersRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = Others)
//     let printBooksRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = PrintBooks)
//     let eBooksRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = EBooks)

//     // let q1Sales = Sales { }
//     0

// let getPacktLibSubscriptions (row: Row) =
//     let q1Copies = row.

let readSheet (sheet2015: Sheet2015) =
    let rows = sheet2015.Data
    let packtLibRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = PacktLib)
    let othersRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = Others)
    let printBooksRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = PrintBooks)
    let eBooksRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = EBooks)

    packtLibRow.Q1