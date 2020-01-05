module OldTypes

open FSharp.Interop.Excel

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
