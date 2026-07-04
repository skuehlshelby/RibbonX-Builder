Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace ComTypes.StdOle

    <ComImport>
    <DefaultMember("Handle")>
    <InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
    <Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB")>
    <ComConversionLoss>
    Public Interface IPictureDisp
        <DispId(0)>
        <ComAliasName("stdole.OLE_HANDLE")>
        ReadOnly Property Handle As Integer

        <DispId(2)>
        <ComAliasName("stdole.OLE_HANDLE")>
        Property hPal As Integer

        <DispId(3)>
        ReadOnly Property Type As Short

        <DispId(4)>
        <ComAliasName("stdole.OLE_XSIZE_HIMETRIC")>
        ReadOnly Property Width As Integer

        <DispId(5)>
        <ComAliasName("stdole.OLE_YSIZE_HIMETRIC")>
        ReadOnly Property Height As Integer

        <DispId(6)>
        <MethodImpl(MethodImplOptions.PreserveSig Or MethodImplOptions.InternalCall)>
        Sub Render(hdc As Integer, x As Integer, y As Integer, cx As Integer, cy As Integer, <ComAliasName("stdole.OLE_XPOS_HIMETRIC")> xSrc As Integer, <ComAliasName("stdole.OLE_YPOS_HIMETRIC")> ySrc As Integer, <ComAliasName("stdole.OLE_XSIZE_HIMETRIC")> cxSrc As Integer, <ComAliasName("stdole.OLE_YSIZE_HIMETRIC")> cySrc As Integer, prcWBounds As IntPtr)

    End Interface

End Namespace
