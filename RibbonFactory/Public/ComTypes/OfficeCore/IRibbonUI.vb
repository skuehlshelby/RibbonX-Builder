Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace ComTypes.Microsoft.Office.Core
    <ComImport>
    <TypeLibType(4160)>
    <Guid("000C03A7-0000-0000-C000-000000000046")>
    Public Interface IRibbonUI
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1)>
        Sub Invalidate()

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(2)>
        Sub InvalidateControl(<[In]> <MarshalAs(UnmanagedType.BStr)> ControlID As String)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(3)>
        Sub InvalidateControlMso(<[In]> <MarshalAs(UnmanagedType.BStr)> ControlID As String)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(4)>
        Sub ActivateTab(<[In]> <MarshalAs(UnmanagedType.BStr)> ControlID As String)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(5)>
        Sub ActivateTabMso(<[In]> <MarshalAs(UnmanagedType.BStr)> ControlID As String)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(6)>
        Sub ActivateTabQ(<[In]> <MarshalAs(UnmanagedType.BStr)> ControlID As String, <[In]> <MarshalAs(UnmanagedType.BStr)> [Namespace] As String)
    End Interface

End Namespace