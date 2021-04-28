Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Controls

Namespace Builders
    
    Public Class CheckBoxBuilder
        Inherits Builder(Of CheckBox)
        Implements ISetEnabled(Of CheckBoxBuilder)
        Implements ISetVisibility(Of CheckBoxBuilder)
        Implements ISetLabelScreenTipAndSuperTip(Of CheckBoxBuilder)
        Implements ISetKeyTip(Of CheckBoxBuilder)
        Implements ISetDescription(Of CheckBoxBuilder)
        Implements ISetAction(Of CheckBoxBuilder)
        
        Public Overrides Function Build() As CheckBox
            Return Build(tag:= Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As CheckBox
            Return New CheckBox(attributes:= Attributes, tag:= tag)
        End Function

        Public Function Enabled() As CheckBoxBuilder Implements ISetEnabled(Of CheckBoxBuilder).Enabled
            AddEnabled(enabled:= True)
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As CheckBoxBuilder Implements ISetEnabled(Of CheckBoxBuilder).Enabled
            AddEnabled(enabled:= True, callback:= callback)
            Return Me
        End Function

        Public Function Disabled() As CheckBoxBuilder Implements ISetEnabled(Of CheckBoxBuilder).Disabled
            AddEnabled(enabled:= False)
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As CheckBoxBuilder Implements ISetEnabled(Of CheckBoxBuilder).Disabled
            AddEnabled(enabled:= False, callback:= callback)
            Return Me
        End Function

        Public Function Visible() As CheckBoxBuilder Implements ISetVisibility(Of CheckBoxBuilder).Visible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As CheckBoxBuilder Implements ISetVisibility(Of CheckBoxBuilder).Visible
            AddVisible(visible:= True, callback:= callback)
            Return Me
        End Function

        Public Function Invisible() As CheckBoxBuilder Implements ISetVisibility(Of CheckBoxBuilder).Invisible
            AddVisible(visible:= False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As CheckBoxBuilder Implements ISetVisibility(Of CheckBoxBuilder).Invisible
            AddVisible(visible:= False, callback:= callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As CheckBoxBuilder Implements ISetLabelScreenTipAndSuperTip(Of CheckBoxBuilder).WithLabel
            AddLabel(label:= label)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label)
            End If
            
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As CheckBoxBuilder Implements ISetLabelScreenTipAndSuperTip(Of CheckBoxBuilder).WithLabel
            AddLabel(label:= label, callback:= callback)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label, callback:= callback)
            End If
            
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As CheckBoxBuilder Implements ISetLabelScreenTipAndSuperTip(Of CheckBoxBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As CheckBoxBuilder Implements ISetLabelScreenTipAndSuperTip(Of CheckBoxBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip, callback:= callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As CheckBoxBuilder Implements ISetLabelScreenTipAndSuperTip(Of CheckBoxBuilder).WithSuperTip
            Throw New NotImplementedException()
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As CheckBoxBuilder Implements ISetLabelScreenTipAndSuperTip(Of CheckBoxBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip, callback:= callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As CheckBoxBuilder Implements ISetKeyTip(Of CheckBoxBuilder).WithKeyTip
            Throw New NotImplementedException()
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As CheckBoxBuilder Implements ISetKeyTip(Of CheckBoxBuilder).WithKeyTip
            Throw New NotImplementedException()
        End Function

        Public Function WithDescription(description As String) As CheckBoxBuilder Implements ISetDescription(Of CheckBoxBuilder).WithDescription
            Throw New NotImplementedException()
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As CheckBoxBuilder Implements ISetDescription(Of CheckBoxBuilder).WithDescription
            Throw New NotImplementedException()
        End Function

        Public Function ThatDoes(callback As OnAction, action As Action) As CheckBoxBuilder Implements ISetAction(Of CheckBoxBuilder).ThatDoes
            Throw New NotImplementedException()
        End Function
    End Class
    
End NameSpace