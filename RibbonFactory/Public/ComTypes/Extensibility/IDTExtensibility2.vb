Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace ComTypes.Extensibility

    <ComImport>
    <Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")>
    <TypeLibType(TypeLibTypeFlags.FDispatchable Or TypeLibTypeFlags.FDual)>
    Public Interface IDTExtensibility2

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1)>
        Sub OnConnection(<[In]> <MarshalAs(UnmanagedType.IDispatch)> Application As Object, <[In]> ConnectMode As ext_ConnectMode, <[In]> <MarshalAs(UnmanagedType.IDispatch)> AddInInst As Object, <[In]> <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(2)>
        Sub OnDisconnection(<[In]> RemoveMode As ext_DisconnectMode, <[In]> <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(3)>
        Sub OnAddInsUpdate(<[In]> <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(4)>
        Sub OnStartupComplete(<[In]> <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(5)>
        Sub OnBegInShutdown(<[In]> <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType:=VarEnum.VT_VARIANT)> ByRef custom As Array)
    End Interface

End Namespace