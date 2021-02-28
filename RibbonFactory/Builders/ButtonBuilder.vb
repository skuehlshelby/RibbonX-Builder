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

        Private ReadOnly _buttonAttributes As AttributeGroup = GetDefaults(Of Button)

        Private _tag As Object = Nothing

        Friend Sub New()

        End Sub

        Public Overrides Function Build() As Button
            Return New Button(_buttonAttributes, _tag)
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithLabel
            _buttonAttributes.Add(New Label(label))
            _buttonAttributes.Add(New ShowLabel(True))

            If copyToScreenTip Then
                _buttonAttributes.Add(New Screentip(label))
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithLabel
            _buttonAttributes.Add(New GetLabel(label, callback))

            If copyToScreenTip Then
                _buttonAttributes.Add(New GetScreentip(label, callback))
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithScreenTip
            _buttonAttributes.Add(New Screentip(screenTip))
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithScreenTip
            _buttonAttributes.Add(New GetScreentip(screenTip, callback))
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithSupertip
            _buttonAttributes.Add(New Supertip(supertip))
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithSupertip
            _buttonAttributes.Add(New GetSupertip(supertip, callback))
            Return Me
        End Function

        Public Function ThatDoes(callback As OnAction, action As Action) As ButtonBuilder
            _buttonAttributes.Add(New Categories.OnAction.OnAction(action, callback))
            Return Me
        End Function
    End Class
End NameSpace