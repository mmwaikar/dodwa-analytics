module Types

open FSharp.Interop.Excel

type QuarterlySales =
    { Q1: double
      Q2: double
      Q3: double
      Q4: double }

type YearlySales =
    { Y2015: QuarterlySales
      Y2016: QuarterlySales
      Y2017: QuarterlySales
      Y2018: QuarterlySales
      Y2019: QuarterlySales }

type SheetPacktLib = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "PacktLib">

type SheetOthers = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "Others">

type SheetPbooks = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "Pbooks">

type SheetEbooks = ExcelFile<"data/books-sold-copy-2-jan-2020.xlsx", "Ebooks">
