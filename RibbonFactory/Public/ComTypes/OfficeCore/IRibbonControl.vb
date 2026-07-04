Imports System.Runtime.InteropServices

Namespace ComTypes.Microsoft.Office.Core

    <ComImport>
    <TypeLibType(4160)>
    <Guid("000C0395-0000-0000-C000-000000000046")>
    Public Interface IRibbonControl

        <DispId(1)>
        ReadOnly Property Id As <MarshalAs(UnmanagedType.BStr)> String


        <DispId(2)>
        ReadOnly Property Context As <MarshalAs(UnmanagedType.IDispatch)> Object


        <DispId(3)>
        ReadOnly Property Tag As <MarshalAs(UnmanagedType.BStr)> String

    End Interface

End Namespace
