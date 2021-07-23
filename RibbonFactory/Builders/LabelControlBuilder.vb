Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders
    Public Class LabelControlBuilder
        Implements IEnabled(Of LabelControlBuilder)
        Implements IInsert(Of LabelControlBuilder)
        Implements IVisible(Of LabelControlBuilder)
        Implements ILabelScreenTipSuperTip(Of LabelControlBuilder)
        Implements IShowLabel(Of LabelControlBuilder)

        Private ReadOnly _builder As ControlBuilder

        Friend Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of LabelControl)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As LabelControl
            Return New LabelControl(_builder.Build(), tag)
        End Function

        Public Function Enabled() As LabelControlBuilder Implements IEnabled(Of LabelControlBuilder).Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements IEnabled(Of LabelControlBuilder).Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As LabelControlBuilder Implements IEnabled(Of LabelControlBuilder).Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements IEnabled(Of LabelControlBuilder).Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As LabelControlBuilder Implements IVisible(Of LabelControlBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements IVisible(Of LabelControlBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As LabelControlBuilder Implements IVisible(Of LabelControlBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements IVisible(Of LabelControlBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As LabelControlBuilder Implements ILabelScreenTipSuperTip(Of LabelControlBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As LabelControlBuilder Implements ILabelScreenTipSuperTip(Of LabelControlBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As LabelControlBuilder Implements ILabelScreenTipSuperTip(Of LabelControlBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As LabelControlBuilder Implements ILabelScreenTipSuperTip(Of LabelControlBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As LabelControlBuilder Implements ILabelScreenTipSuperTip(Of LabelControlBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As LabelControlBuilder Implements ILabelScreenTipSuperTip(Of LabelControlBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function ShowLabel() As LabelControlBuilder Implements IShowLabel(Of LabelControlBuilder).ShowLabel
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements IShowLabel(Of LabelControlBuilder).ShowLabel
            _builder.ShowLabel(callback)
            Return Me
        End Function

        Public Function HideLabel() As LabelControlBuilder Implements IShowLabel(Of LabelControlBuilder).HideLabel
            _builder.HideLabel()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements IShowLabel(Of LabelControlBuilder).HideLabel
            _builder.HideLabel(callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As LabelControlBuilder Implements IInsert(Of LabelControlBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As LabelControlBuilder Implements IInsert(Of LabelControlBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As LabelControlBuilder Implements IInsert(Of LabelControlBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As LabelControlBuilder Implements IInsert(Of LabelControlBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

    End Class

End NameSpace