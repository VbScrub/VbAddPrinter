
Imports System.Runtime.InteropServices

Module MainModule

    Sub Main()
        Try
            If My.Application.CommandLineArgs.Count = 0 OrElse My.Application.CommandLineArgs(0) = "-h" OrElse My.Application.CommandLineArgs(0) = "/?" Then
                Console.WriteLine("Valid arguments: " & Environment.NewLine &
                              "VbAddPrinter.exe add [printer_name] [driver_name] [port_name]""" & Environment.NewLine &
                              "Or" & Environment.NewLine &
                              "VbAddPrinter.exe enum" & Environment.NewLine &
                              "Or" & Environment.NewLine &
                              "VbAddPrinter.exe addunc [\\server\printer]")
                Return
            End If
            Dim FirstArgument As String = My.Application.CommandLineArgs(0).ToLower
            If FirstArgument = "add" Then
                Dim PrinterInfo As New WinApi.PRINTER_INFO_2
                PrinterInfo.pPrinterName = My.Application.CommandLineArgs(1)
                PrinterInfo.pDriverName = My.Application.CommandLineArgs(2)
                PrinterInfo.pPortName = My.Application.CommandLineArgs(3)
                Dim printerHandle As IntPtr = WinApi.AddPrinter(Nothing, 2, PrinterInfo)
                If Not printerHandle = IntPtr.Zero Then
                    Console.Write("Printer successfully added")
                    WinApi.ClosePrinter(printerHandle)
                Else
                    Throw New ComponentModel.Win32Exception
                End If
            ElseIf FirstArgument = "addunc" Then
                Console.WriteLine("go")
                Threading.Thread.Sleep(15000)
                If Not WinApi.AddPrinterConnection(My.Application.CommandLineArgs(1)) Then
                    Throw New ComponentModel.Win32Exception
                End If
            ElseIf FirstArgument = "enum" Then
                    Dim NumberOfItems As UInteger
                Dim PrintersPtr As IntPtr
                Dim StepSize As Integer = Marshal.SizeOf(GetType(WinApi.PRINTER_INFO_1))
                Dim RequiredBufferSize As UInteger
                WinApi.EnumPrinters(WinApi.PRINTER_ENUM_LOCAL, Nothing, 1, PrintersPtr, 0, RequiredBufferSize, NumberOfItems)
                PrintersPtr = Marshal.AllocHGlobal(CInt(RequiredBufferSize))
                If WinApi.EnumPrinters(WinApi.PRINTER_ENUM_LOCAL, Nothing, 1, PrintersPtr, RequiredBufferSize, 0, NumberOfItems) Then
                    For i As Integer = 0 To CInt(NumberOfItems - 1)
                        Dim CurrentPrinter As WinApi.PRINTER_INFO_1 = Marshal.PtrToStructure(Of WinApi.PRINTER_INFO_1)(PrintersPtr)
                        Console.WriteLine("Name: " & CurrentPrinter.pName)
                        PrintersPtr = IntPtr.Add(PrintersPtr, StepSize)
                    Next
                    Marshal.FreeHGlobal(PrintersPtr)
                Else
                    Throw New ComponentModel.Win32Exception
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

End Module
