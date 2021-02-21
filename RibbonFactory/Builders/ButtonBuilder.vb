Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Supertip

Namespace Builders
    Public NotInheritable Class ButtonBuilder
        Inherits Builder(Of Button)
        Implements IDescribable(Of ButtonBuilder)

        Private ReadOnly ButtonAttributes As AttributeGroup = New AttributeGroup()

        Private Tag As Object = Nothing
        Friend Sub New()

        End Sub

        Public Overrides Function Build() As Button
            Return New Button(ButtonAttributes, Tag)
        End Function

        Public Function WithLabel(Label As String, Optional CopyToScreentip As Boolean = True) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithLabel
            ButtonAttributes.Add(New Label(Label))

            If CopyToScreentip Then
                ButtonAttributes.Add(New Screentip(Label))
            End If

            Return Me
        End Function

        Public Function WithLabel(Label As String, Callback As FromControl(Of String), Optional CopyToScreentip As Boolean = True) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithLabel
            ButtonAttributes.Add(New GetLabel(Label, Callback))

            If CopyToScreentip Then
                ButtonAttributes.Add(New GetScreentip(Label, Callback))
            End If

            Return Me
        End Function

        Public Function WithScreenTip(ScreenTip As String) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithScreenTip
            ButtonAttributes.Add(New Screentip(ScreenTip))
            Return Me
        End Function

        Public Function WithScreenTip(ScreenTip As String, Callback As FromControl(Of String)) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithScreenTip
            ButtonAttributes.Add(New GetScreentip(ScreenTip, Callback))
            Return Me
        End Function

        Public Function WithSupertip(Supertip As String) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithSupertip
            ButtonAttributes.Add(New Supertip(Supertip))
            Return Me
        End Function

        Public Function WithSupertip(Supertip As String, Callback As FromControl(Of String)) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithSupertip
            ButtonAttributes.Add(New GetSupertip(Supertip, Callback))
            Return Me
        End Function

        Public Function ThatDoes(Callback As OnAction, Action As Action) As ButtonBuilder
            ButtonAttributes.Add(New RibbonAttributes.Categories.OnAction.OnAction(Action, Callback))
        End Function
    End Class
End NameSpace