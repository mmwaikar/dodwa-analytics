// Learn more about F# at http://fsharp.org

open System
open Types
open ExcelUtil
open DrawViz

[<EntryPoint>]
let main argv =
    printfn "Hello World from F#!"

    let packtLibData = getPacktLibData
    let othersData = getOthersData
    let pBooksData = getPBooksData
    let eBooksData = getEBooksData

    // printfn "PacktLib Sales: %A" packtLibData
    // printfn "Others Sales: %A" othersData
    // printfn "Print books Sales: %A" pBooksData
    // printfn "Ebooks Sales: %A" eBooksData

    // individual charts
    // (drawQuarterWiseChart (packtLibData, "PacktLib Subscriptions")).Show()
    // (drawQuarterWiseChart (othersData, "Other Subscriptions")).Show()
    // (drawQuarterWiseChart (pBooksData, "Print books")).Show()
    // (drawQuarterWiseChart (eBooksData, "Ebooks")).Show()

    // the below two are similar (the first is , the second one is line)
    (drawYearlyChart (packtLibData, othersData, pBooksData, eBooksData)).Show()
    (drawLineChart (packtLibData, othersData, pBooksData, eBooksData)).Show()

    // this is the consolidation of the drawYearlyChart (above)
    (drawYearWiseStackedChart (packtLibData, othersData, pBooksData, eBooksData)).Show()
    // (drawItemWiseStackedChart (packtLibData, othersData, pBooksData, eBooksData)).Show()
    (drawTotalChart (packtLibData, othersData, pBooksData, eBooksData)).Show()

    0 // return an integer exit code
