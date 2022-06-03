Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace ComTypes.StdOle

    <ComImport>
    <ComConversionLoss>
    <Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPicture

        <DispId(1610678272)>
        <ComAliasName("stdole.OLE_HANDLE")>
        ReadOnly Property Handle As Integer

        <DispId(1610678273)>
        <ComAliasName("stdole.OLE_HANDLE")>
        ReadOnly Property hPal As Integer

        <DispId(1610678274)>
        ReadOnly Property Type As Short

        <DispId(1610678275)>
        <ComAliasName("stdole.OLE_XSIZE_HIMETRIC")>
        ReadOnly Property Width As Integer

        <DispId(1610678276)>
        <ComAliasName("stdole.OLE_YSIZE_HIMETRIC")>
        ReadOnly Property Height As Integer

        <DispId(1610678279)>
        ReadOnly Property CurDC As Integer

        <DispId(1610678281)>
        Property KeepOriginalFormat As Boolean

        <DispId(1610678285)>
        ReadOnly Property Attributes As Integer

        <MethodImpl(MethodImplOptions.InternalCall)>
        Sub Render(<[In]> hdc As Integer, <[In]> x As Integer, <[In]> y As Integer, <[In]> cx As Integer, <[In]> cy As Integer, <[In]> <ComAliasName("stdole.OLE_XPOS_HIMETRIC")> xSrc As Integer, <[In]> <ComAliasName("stdole.OLE_YPOS_HIMETRIC")> ySrc As Integer, <[In]> <ComAliasName("stdole.OLE_XSIZE_HIMETRIC")> cxSrc As Integer, <[In]> <ComAliasName("stdole.OLE_YSIZE_HIMETRIC")> cySrc As Integer, <[In]> prcWBounds As IntPtr)

        <MethodImpl(MethodImplOptions.InternalCall)>
        Sub SelectPicture(<[In]> hdcIn As Integer, <Out> ByRef phdcOut As Integer, <ComAliasName("stdole.OLE_HANDLE")> <Out> ByRef phbmpOut As Integer)

        <MethodImpl(MethodImplOptions.InternalCall)>
        Sub PictureChanged()

        <MethodImpl(MethodImplOptions.InternalCall)>
        Sub SaveAsFile(<[In]> pstm As IntPtr, <[In]> fSaveMemCopy As Boolean, <Out> ByRef pcbSize As Integer)

        <MethodImpl(MethodImplOptions.InternalCall)>
        Sub SetHdc(<[In]> <ComAliasName("stdole.OLE_HANDLE")> hdc As Integer)

    End Interface

End Namespace