Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders
    
    Public Class CheckBoxBuilder
        Implements IEnabled(Of CheckBoxBuilder)
        Implements IVisible(Of CheckBoxBuilder)
        Implements ILabelScreenTipSuperTip(Of CheckBoxBuilder)
        Implements IInsert(Of CheckBoxBuilder)
        Implements IKeyTip(Of CheckBoxBuilder)
        Implements IDescription(Of CheckBoxBuilder)
        Implements IOnActionToggle(Of CheckBoxBuilder)

        Private ReadOnly _builder As ControlBuilder

        Friend Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of CheckBox)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As CheckBox
            Return New CheckBox(_builder.Build(), tag:= tag)
        End Function

        Public Function Enabled() As CheckBoxBuilder Implements IEnabled(Of CheckBoxBuilder).Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As CheckBoxBuilder Implements IEnabled(Of CheckBoxBuilder).Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As CheckBoxBuilder Implements IEnabled(Of CheckBoxBuilder).Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As CheckBoxBuilder Implements IEnabled(Of CheckBoxBuilder).Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As CheckBoxBuilder Implements IVisible(Of CheckBoxBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As CheckBoxBuilder Implements IVisible(Of CheckBoxBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As CheckBoxBuilder Implements IVisible(Of CheckBoxBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As CheckBoxBuilder Implements IVisible(Of CheckBoxBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As CheckBoxBuilder Implements ILabelScreenTipSuperTip(Of CheckBoxBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As CheckBoxBuilder Implements ILabelScreenTipSuperTip(Of CheckBoxBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As CheckBoxBuilder Implements ILabelScreenTipSuperTip(Of CheckBoxBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As CheckBoxBuilder Implements ILabelScreenTipSuperTip(Of CheckBoxBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As CheckBoxBuilder Implements ILabelScreenTipSuperTip(Of CheckBoxBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As CheckBoxBuilder Implements ILabelScreenTipSuperTip(Of CheckBoxBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function WithDescription(description As String) As CheckBoxBuilder Implements IDescription(Of CheckBoxBuilder).WithDescription
            _builder.WithDescription(description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As CheckBoxBuilder Implements IDescription(Of CheckBoxBuilder).WithDescription
            _builder.WithDescription(description, callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As CheckBoxBuilder Implements IKeyTip(Of CheckBoxBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As CheckBoxBuilder Implements IKeyTip(Of CheckBoxBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function ThatDoes(action As Action, callback As ButtonPressed) As CheckBoxBuilder Implements IOnActionToggle(Of CheckBoxBuilder).ThatDoes
            _builder.ThatDoes(action, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As CheckBoxBuilder Implements IInsert(Of CheckBoxBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As CheckBoxBuilder Implements IInsert(Of CheckBoxBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As CheckBoxBuilder Implements IInsert(Of CheckBoxBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As CheckBoxBuilder Implements IInsert(Of CheckBoxBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

    End Class
    
End NameSpace