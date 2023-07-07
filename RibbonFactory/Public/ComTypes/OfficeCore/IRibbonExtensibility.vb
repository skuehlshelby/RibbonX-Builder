Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace ComTypes.Microsoft.Office.Core

    <ComImport>
    <Guid("000C0396-0000-0000-C000-000000000046")>
    <TypeLibType(4160)>
    Public Interface IRibbonExtensibility

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1)>
        Function GetCustomUI(<[In]> <MarshalAs(UnmanagedType.BStr)> RibbonID As String) As <MarshalAs(UnmanagedType.BStr)> String

    End Interface

End Namespace
