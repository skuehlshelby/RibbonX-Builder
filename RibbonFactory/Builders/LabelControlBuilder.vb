Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders
    Public Class LabelControlBuilder
        Implements IBuilder(Of LabelControl)
        Implements IID(Of LabelControlBuilder)
        Implements IEnabled(Of LabelControlBuilder)
        Implements IInsert(Of LabelControlBuilder)
        Implements IVisible(Of LabelControlBuilder)
        Implements ILabelScreenTipSuperTip(Of LabelControlBuilder)
        Implements IShowLabel(Of LabelControlBuilder)

        Private ReadOnly _builder As ControlBuilder

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New Defaults(Of LabelControl)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As LabelControl)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(template)
            attributeGroupBuilder.AddID(IdManager.GetID(Of LabelControl)())
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As RibbonElement)
            Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

            If defaultProvider IsNot Nothing Then
                Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
                Dim labelControlAttributes As AttributeSet = New Defaults(Of LabelControl)().GetDefaults()
                Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) labelControlAttributes.Contains(a)))
                Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
                attributeGroupBuilder.AddID(IdManager.GetID(Of LabelControl)())
                _builder = New ControlBuilder(attributeGroupBuilder)
            Else
                Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(LabelControl)}'")
            End If
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As LabelControl Implements IBuilder(Of LabelControl).Build
            Return New LabelControl(_builder.Build(), tag)
        End Function

        Public Function WithId(id As String) As LabelControlBuilder Implements IID(Of LabelControlBuilder).WithId
            _builder.WithId(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As LabelControlBuilder Implements IID(Of LabelControlBuilder).WithIdQ
            _builder.WithId([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As LabelControlBuilder Implements IID(Of LabelControlBuilder).WithIdMso
            _builder.WithId(mso)
            Return Me
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
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As LabelControlBuilder Implements ILabelScreenTipSuperTip(Of LabelControlBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            _builder.ShowLabel()
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