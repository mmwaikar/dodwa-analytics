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

    (drawChart (packtLibData, "PacktLib Subscriptions")).Show()
    (drawChart (othersData, "Other Subscriptions")).Show()
    (drawChart (pBooksData, "Print books")).Show()
    (drawChart (eBooksData, "Ebooks")).Show()

    0 // return an integer exit code
