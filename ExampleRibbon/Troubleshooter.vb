Imports System.Runtime.ExceptionServices
Imports System.Text
Imports Extensibility

Friend Class Troubleshooter
    Implements IDTExtensibility2

    Public Sub New()
        With AppDomain.CurrentDomain
            AddHandler .FirstChanceException, AddressOf OnExceptionThrown
            AddHandler .UnhandledException, AddressOf OnUnhandledException
        End With
    End Sub

    Private Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        MessageBox("Connected!")
    End Sub

    Private Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        MessageBox("Disconnected!")
    End Sub

    Private Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
        MessageBox("Add-Ins Updated!")
    End Sub

    Private Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
        MessageBox("Startup Complete!")
    End Sub

    Private Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
        MessageBox("Shutdown Starting!")
    End Sub

    Private Shared Sub MessageBox(prompt As String)
        MsgBox(prompt, MsgBoxStyle.OkOnly, My.Application.Info.Title)
    End Sub

    Private Shared Sub OnExceptionThrown(sender As Object, e As FirstChanceExceptionEventArgs)
        MessageBox(FormatException(e.Exception))
    End Sub

    Private Shared Sub OnUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        MessageBox(FormatException(e.ExceptionObject))
    End Sub

    Private Shared Function FormatException(e As Exception) As String
        With New StringBuilder()
            .AppendLine($"A '{e.GetType().Name}' was thrown in method '{e.TargetSite.Name}'.")
            .AppendLine()
            .AppendLine("---------Message----------")
            .AppendLine(e.Message)
            .AppendLine()
            .AppendLine("-------Stack Trace--------")
            .AppendLine(e.StackTrace)

            If e.Data IsNot Nothing AndAlso e.Data.Count > 0 Then
                .AppendLine()
                .AppendLine("----------Data------------")
                For Each datum As Object In e.Data
                    .AppendLine(datum.ToString())
                Next
            End If

            Return .ToString()
        End With
    End Function
End Class