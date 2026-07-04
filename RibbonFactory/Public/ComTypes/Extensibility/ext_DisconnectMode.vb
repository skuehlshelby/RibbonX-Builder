Imports System.Runtime.InteropServices

Namespace ComTypes.Extensibility

    <Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")>
    Public Enum ext_DisconnectMode
        ext_dm_HostShutdown
        ext_dm_UserClosed
        ext_dm_UISetupComplete
        ext_dm_SolutionClosed
    End Enum

End Namespace