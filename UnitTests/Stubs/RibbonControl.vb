Imports Microsoft.Office.Core

Namespace Stubs

    Public Class RibbonControl
        Implements IRibbonControl

        Public Sub New(controlID As String, context As Object, tag As String)
            Id = controlID
            Me.Context = context
            Me.Tag = tag
        End Sub

        Public ReadOnly Property Id As String Implements IRibbonControl.Id

        Public ReadOnly Property Context As Object Implements IRibbonControl.Context

        Public ReadOnly Property Tag As String Implements IRibbonControl.Tag
    End Class
End NameSpace