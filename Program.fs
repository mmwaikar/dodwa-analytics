// Learn more about F# at http://fsharp.org

open System
open ExcelUtil

[<EntryPoint>]
let main argv =
    printfn "Hello World from F#!"
    let q1 = readSheet (Sheet2015())
    printfn "%A" q1
    0 // return an integer exit code
