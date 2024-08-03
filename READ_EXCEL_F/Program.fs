open MODELS 
open System
open System.Runtime.InteropServices
open System.Reflection
open System
open System.IO
open System.Threading

let executeMacro () =
    let excelType = Type.GetTypeFromProgID("Excel.Application")
    let excelApp = Activator.CreateInstance(excelType)
    let excelSerialToDate (serialNumber: int) =
    // Excel considera el 1 de enero de 1900 como el día 1
        let baseDate = DateTime(1900, 1, 1)
        baseDate.AddDays(float serialNumber - 1.0) // Restar 1 porque Excel empieza desde 1

    try
        excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, [| box true |]) |> ignore
        
        let workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null)
        let fileName = "Calculadora de préstamos simple y tabla de amortización.xlsx"
        let filePath = $"C:\{fileName}";
        let workbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, [| $"{filePath}" |])
        
        try
            printf "Ingrese monto prestamo: "
            let montoPrestamo = Console.ReadLine() |> Convert.ToDouble
            
            printf "Ingrese tasa interes: "
            let tasaInteres = Console.ReadLine()
            
            printf "Ingrese periodo (meses): "
            let periodo = Console.ReadLine() |> Convert.ToInt32
            
            let worksheets = workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, workbook, null)
            let datosSheet = worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, [| box "Calculadora de préstamos" |])
            
            let cellD3 = datosSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, datosSheet, [| box 3; box 4 |])
            cellD3.GetType().InvokeMember("Value2", BindingFlags.SetProperty, null, cellD3, [| box montoPrestamo |]) |> ignore
            printfn "Monto préstamo '%f' insertado en D3." montoPrestamo
            
            let cellD4 = datosSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, datosSheet, [| box 4; box 4 |])
            cellD4.GetType().InvokeMember("Value2", BindingFlags.SetProperty, null, cellD4, [| box tasaInteres |]) |> ignore
            printfn "Tasa interés '%s' insertada en D4." tasaInteres 
            
            let cellD5 = datosSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, datosSheet, [| box 5; box 4 |])
            cellD5.GetType().InvokeMember("Value2", BindingFlags.SetProperty, null, cellD5, [| box periodo |]) |> ignore
            printfn "Periodo '%d' insertado en D5." periodo
          
            let imprimirSheet = worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, [| box "Calculadora de préstamos" |])
            
            Thread.Sleep(10000)
            let rec leerDatos fila acumulado =
                let cellNumero = imprimirSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, imprimirSheet, [| box fila; box 2 |])
                let cellNumeroValue = cellNumero.GetType().InvokeMember("Value2", BindingFlags.GetProperty, null, cellNumero, null) |> string

                if String.IsNullOrWhiteSpace(cellNumeroValue) then
                    acumulado
                else
                    let cellFecha = imprimirSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, imprimirSheet, [| box fila; box 3 |])
                    let cellSaldo = imprimirSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, imprimirSheet, [| box fila; box 4 |])
                    let cellInteres = imprimirSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, imprimirSheet, [| box fila; box 5 |])
                    let cellSegDes = imprimirSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, imprimirSheet, [| box fila; box 6 |])
                    let cellSegBien = imprimirSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, imprimirSheet, [| box fila; box 7 |])
                    let cellAmort = imprimirSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, imprimirSheet, [| box fila; box 8 |])
                    let cellCuota = imprimirSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, imprimirSheet, [| box fila; box 9 |])
                    
                    let cellValue cell = cell.GetType().InvokeMember("Value2", BindingFlags.GetProperty, null, cell, null) |> string
                    
                    let nuevaFila = { 
                        Numero = cellValue cellNumero
                        Fecha = cellValue cellFecha
                        Saldo = cellValue cellSaldo
                        Interes = cellValue cellInteres
                        SegDes = cellValue cellSegDes
                        SegBien = cellValue cellSegBien
                        Amort = cellValue cellAmort
                    }
                    
                    leerDatos (fila + 1) (acumulado @ [nuevaFila])
            
            let datos = leerDatos 14 []
            
            datos |> List.iter (fun fila -> 
                printfn "Nº: %s, Fecha de pago: %s, Salgo inicial: %s, Pago: %s, Principal: %s, Interés: %s, Saldo final: %s" 
                        fila.Numero fila.Fecha fila.Saldo fila.Interes fila.SegDes fila.SegBien fila.Amort )
        
        finally
            workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, [| box false |]) |> ignore
            Marshal.ReleaseComObject(workbook) |> ignore

    finally
        excelType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null) |> ignore
        Marshal.ReleaseComObject(excelApp) |> ignore

executeMacro ()
