Imports System.Runtime.InteropServices

Public Class WinApi
    <StructLayout(LayoutKind.Sequential)>
    Public Structure PRINTER_INFO_2
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pServerName As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pPrinterName As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pShareName As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pPortName As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pDriverName As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pComment As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pLocation As String
        Public pDevMode As IntPtr
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pSepFile As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pPrintProcessor As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pDatatype As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pParameters As String
        Public pSecurityDescriptor As IntPtr
        Public Attributes As UInteger
        Public Priority As UInteger
        Public DefaultPriority As UInteger
        Public StartTime As UInteger
        Public UntilTime As UInteger
        Public Status As UInteger
        Public cJobs As UInteger
        Public AveragePPM As UInteger
    End Structure

    <StructLayoutAttribute(LayoutKind.Sequential)>
    Public Structure PRINTER_INFO_1
        Public Flags As UInteger
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pDescription As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pName As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pComment As String
    End Structure

    Public Const PRINTER_ENUM_LOCAL As Integer = 2


    <DllImport("Winspool.drv", EntryPoint:="AddPrinterW", SetLastError:=True)>
    Public Shared Function AddPrinter(<InAttribute(), MarshalAs(UnmanagedType.LPWStr)> ByVal pName As String, ByVal Level As UInteger,
                                      <InAttribute()> ByRef pPrinter As PRINTER_INFO_2) As System.IntPtr
    End Function

    <DllImport("Winspool.drv", EntryPoint:="ClosePrinter", SetLastError:=True)>
    Public Shared Function ClosePrinter(<InAttribute()> ByVal Handle As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("Winspool.drv", EntryPoint:="EnumPrintersW", SetLastError:=True)>
    Public Shared Function EnumPrinters(ByVal Flags As UInteger, <InAttribute(), MarshalAs(UnmanagedType.LPWStr)> ByVal Name As String,
                                         ByVal Level As UInteger, ByVal pPrinterEnum As IntPtr, ByVal cbBuf As UInteger,
                                         <OutAttribute()> ByRef pcbNeeded As UInteger,
                                         <OutAttribute()> ByRef pcReturned As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("Winspool.drv", EntryPoint:="AddPrinterConnectionW", SetLastError:=True)>
    Public Shared Function AddPrinterConnection(<InAttribute(), MarshalAs(UnmanagedType.LPWStr)> ByVal Name As String) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function



End Class
