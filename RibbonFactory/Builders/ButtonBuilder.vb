Imports RibbonFactory.CallbackSignatures

Public Class ButtonBuilder
    Inherits Builder(Of Button)
    Implements IDescribable(Of ButtonBuilder)

    Private ButtonAttributes As ControlPropertyList = AttributeProvider.GetDefaults(Of Button)()
    Private Tag As Object = Nothing
    Friend Sub New()

    End Sub

    Public Overrides Function Build() As Button
        Return New Button(ButtonAttributes, Tag)
    End Function

    Public Function WithLabel(Label As String, Optional CopyToScreentip As Boolean = True) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithLabel
        ButtonAttributes.Add(New ControlProperty(Of String)(Attributes.Label.Label, Label))

        If CopyToScreentip Then
            ButtonAttributes.Add(New ControlProperty(Of String)(Attributes.Screentip.Screentip, Label))
        End If

        Return Me
    End Function

    Public Function WithLabel(Label As String, Callback As FromControl(Of String), Optional CopyToScreentip As Boolean = True) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithLabel
        ButtonAttributes.Add(New DynamicControlProperty(Of String)(Attributes.Label.GetLabel, Callback.Method.Name, Label))

        If CopyToScreentip Then
            ButtonAttributes.Add(New DynamicControlProperty(Of String)(Attributes.Screentip.GetScreentip, Callback.Method.Name, Label))
        End If

        Return Me
    End Function

    Public Function WithScreenTip(ScreenTip As String) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithScreenTip
        ButtonAttributes.Add(New ControlProperty(Of String)(Attributes.Screentip.Screentip, ScreenTip))
        Return Me
    End Function

    Public Function WithScreenTip(ScreenTip As String, Callback As FromControl(Of String)) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithScreenTip
        ButtonAttributes.Add(New DynamicControlProperty(Of String)(Attributes.Screentip.GetScreentip, Callback.Method.Name, ScreenTip))
        Return Me
    End Function

    Public Function WithSupertip(Supertip As String) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithSupertip
        ButtonAttributes.Add(New ControlProperty(Of String)(Attributes.Supertip.Supertip, Supertip))
        Return Me
    End Function

    Public Function WithSupertip(Supertip As String, Callback As FromControl(Of String)) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithSupertip
        ButtonAttributes.Add(New DynamicControlProperty(Of String)(Attributes.Supertip.GetSupertip, Callback.Method.Name, Supertip))
        Return Me
    End Function

    Public Function ThatDoes(ByVal Callback As OnAction) As ButtonBuilder

    End Function

    Public Function ThatDoes(ByVal Callback As OnAction, ByVal Action As Action) As ButtonBuilder

    End Function
End Class
