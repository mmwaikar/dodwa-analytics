module OldExcelUtil

open FSharp.Interop.Excel
open Types

let PacktLib = "Subscription PacktLib"
let Others = "Subscription Others"
let PrintBooks = "Book copies Print"
let EBooks = "Book copies Ebook"

let PacktLibRowIndex = 1
let OthersRowIndex = 2
let PrintBooksRowIndex = 3
let EBooksRowIndex = 4

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
    let data = sheet2015.Data
    let rows = List.ofSeq (data)
    // let packtLibRow = rows |> List.find (fun row -> row.``Till Date ->``.ToString() = PacktLib)
    // let othersRow =
    //     rows |> Seq.find (fun row -> not (isNull (row.``Till Date ->``)) && row.``Till Date ->``.ToString() = Others)
    // let printBooksRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = PrintBooks)
    // let eBooksRow = rows |> Seq.find (fun row -> row.``Till Date ->``.ToString() = EBooks)

    let othersRow = rows.Item(OthersRowIndex)
    let others = othersRow.Q1
    printfn "Others in Q1 2015: %s" others
    // printfn "PacktLib in Q1 2015: %A" othersRow.Q1
    printfn "Total rows: %d" rows.Length
    rows.Item(0)
 