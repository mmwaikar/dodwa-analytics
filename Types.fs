module Types

open FSharp.Interop.Excel

type QuarterlySales =
    { Q1: double
      Q2: double
      Q3: double
      Q4: double }

type QuarterWiseYearlySales =
    { Y2015: QuarterlySales
      Y2016: QuarterlySales
      Y2017: QuarterlySales
      Y2018: QuarterlySales
      Y2019: QuarterlySales
      Y2020: QuarterlySales }

type YearlySales =
    { Y2015: double
      Y2016: double
      Y2017: double
      Y2018: double
      Y2019: double
      Y2020: double }

type SheetPacktLib = ExcelFile<"data/books-sold-as-on-20-mar-2021.xlsx", "PacktLib">

type SheetOthers = ExcelFile<"data/books-sold-as-on-20-mar-2021.xlsx", "Others">

type SheetPbooks = ExcelFile<"data/books-sold-as-on-20-mar-2021.xlsx", "Pbooks">

type SheetEbooks = ExcelFile<"data/books-sold-as-on-20-mar-2021.xlsx", "Ebooks">
