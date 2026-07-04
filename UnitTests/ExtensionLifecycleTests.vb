Imports RibbonX
Imports RibbonX.Controls
Imports RibbonX.ComTypes.Extensibility

''' <summary>
''' Exercises the <see cref="CustomRibbonBase"/> IDTExtensibility2 fan-out: the base class
''' multiplexes each Office add-in lifecycle callback to every composed extension, isolates
''' a throwing extension so the others still run, and still performs its own teardown.
''' </summary>
<TestClass()>
Public Class ExtensionLifecycleTests

    <TestMethod()>
    Public Sub OnConnection_FansOutToEveryExtension()
        Dim first As New RecordingExtension()
        Dim second As New RecordingExtension()
        Dim ribbon As New HostedRibbon(first, second)

        Dim custom As Array = Nothing
        ribbon.OnConnection(New Object(), ext_ConnectMode.ext_cm_Startup, New Object(), custom)

        Assert.AreEqual(1, first.ConnectionCount)
        Assert.AreEqual(1, second.ConnectionCount)
    End Sub

    <TestMethod()>
    Public Sub OnConnection_CapturesHostApplication()
        Dim host As Object = New Object()
        Dim ribbon As New HostedRibbon()

        Dim custom As Array = Nothing
        ribbon.OnConnection(host, ext_ConnectMode.ext_cm_Startup, New Object(), custom)

        Assert.AreSame(host, ribbon.CapturedHostApp)
    End Sub

    <TestMethod()>
    Public Sub OnConnection_ThrowingExtensionDoesNotStopTheOthers()
        Dim before As New RecordingExtension()
        Dim after As New RecordingExtension()
        Dim ribbon As New HostedRibbon(before, New ThrowingExtension(), after)

        Dim custom As Array = Nothing
        ribbon.OnConnection(New Object(), ext_ConnectMode.ext_cm_Startup, New Object(), custom)

        Assert.AreEqual(1, before.ConnectionCount, "Extensions before the failure should still be called.")
        Assert.AreEqual(1, after.ConnectionCount, "A failure must not prevent later extensions from connecting.")
    End Sub

    <TestMethod()>
    Public Sub OnDisconnection_ClearsHostApp_EvenWhenAnExtensionThrows()
        Dim ribbon As New HostedRibbon(New ThrowingExtension())

        Dim custom As Array = Nothing
        ribbon.OnConnection(New Object(), ext_ConnectMode.ext_cm_Startup, New Object(), custom)
        Assert.IsNotNull(ribbon.CapturedHostApp)

        ribbon.OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown, custom)

        Assert.IsNull(ribbon.CapturedHostApp, "Teardown after the fan-out must run even if an extension throws.")
    End Sub

    <TestMethod()>
    Public Sub EveryLifecycleCallbackIsForwarded()
        Dim ext As New RecordingExtension()
        Dim ribbon As New HostedRibbon(ext)

        Dim custom As Array = Nothing
        ribbon.OnConnection(New Object(), ext_ConnectMode.ext_cm_Startup, New Object(), custom)
        ribbon.OnAddInsUpdate(custom)
        ribbon.OnStartupComplete(custom)
        ribbon.OnBeginShutdown(custom)
        ribbon.OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown, custom)

        Assert.AreEqual(1, ext.ConnectionCount)
        Assert.AreEqual(1, ext.AddInsUpdateCount)
        Assert.AreEqual(1, ext.StartupCompleteCount)
        Assert.AreEqual(1, ext.BeginShutdownCount)
        Assert.AreEqual(1, ext.DisconnectionCount)
    End Sub

#Region "Test Doubles"

    ''' <summary>A concrete ribbon that lets the tests reach the protected HostApp seam.</summary>
    Private NotInheritable Class HostedRibbon
        Inherits CustomRibbonBase

        Public Sub New(ParamArray extensions() As IDTExtensibility2)
            MyBase.New(extensions)
        End Sub

        Public ReadOnly Property CapturedHostApp As Object
            Get
                Return HostApp
            End Get
        End Property

        Protected Overrides Function BuildRibbon() As IRibbon
            ' Not exercised by these tests; the lifecycle seam is independent of ribbon XML.
            Throw New NotSupportedException()
        End Function
    End Class

    ''' <summary>An extension that records how often each lifecycle callback fired.</summary>
    Private NotInheritable Class RecordingExtension
        Implements IDTExtensibility2

        Public Property ConnectionCount As Integer
        Public Property DisconnectionCount As Integer
        Public Property AddInsUpdateCount As Integer
        Public Property StartupCompleteCount As Integer
        Public Property BeginShutdownCount As Integer

        Public Sub OnConnection(application As Object, connectMode As ext_ConnectMode, addInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
            ConnectionCount += 1
        End Sub

        Public Sub OnDisconnection(removeMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
            DisconnectionCount += 1
        End Sub

        Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
            AddInsUpdateCount += 1
        End Sub

        Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
            StartupCompleteCount += 1
        End Sub

        Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBegInShutdown
            BeginShutdownCount += 1
        End Sub
    End Class

    ''' <summary>An extension that faults on every callback, to verify failure isolation.</summary>
    Private NotInheritable Class ThrowingExtension
        Implements IDTExtensibility2

        Public Sub OnConnection(application As Object, connectMode As ext_ConnectMode, addInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
            Throw New InvalidOperationException("boom")
        End Sub

        Public Sub OnDisconnection(removeMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
            Throw New InvalidOperationException("boom")
        End Sub

        Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
            Throw New InvalidOperationException("boom")
        End Sub

        Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
            Throw New InvalidOperationException("boom")
        End Sub

        Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBegInShutdown
            Throw New InvalidOperationException("boom")
        End Sub
    End Class

#End Region

End Class
