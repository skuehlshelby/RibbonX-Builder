Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders

    Public Class CheckBoxBuilder
        Implements IBuilder(Of CheckBox)
        Implements IID(Of CheckBoxBuilder)
        Implements IEnabled(Of CheckBoxBuilder)
        Implements IVisible(Of CheckBoxBuilder)
        Implements ILabelScreenTipSuperTip(Of CheckBoxBuilder)
        Implements IInsert(Of CheckBoxBuilder)
        Implements IKeyTip(Of CheckBoxBuilder)
        Implements IDescription(Of CheckBoxBuilder)

        Private ReadOnly _builder As ControlBuilder

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of CheckBox)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As CheckBox)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(template)
            attributeGroupBuilder.AddID(IdManager.GetID(Of CheckBox)())
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As Object)
            Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

            If defaultProvider IsNot Nothing Then
                Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
                Dim checkBoxAttributes As AttributeSet = New DefaultProvider(Of CheckBox)().GetDefaults()
                Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) checkBoxAttributes.Contains(a)))
                Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
                attributeGroupBuilder.AddID(IdManager.GetID(Of CheckBox)())
                _builder = New ControlBuilder(attributeGroupBuilder)
            Else
                Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(CheckBox)}'")
            End If
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As CheckBox Implements IBuilder(Of CheckBox).Build
            Return New CheckBox(_builder.Build(), tag:=tag)
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

        Public Function Checked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As CheckBoxBuilder
            _builder.GetPressedFrom(True, getChecked, setChecked)
            Return Me
        End Function

        Public Function Checked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed, action As Action(Of CheckBox)) As CheckBoxBuilder
            _builder.GetPressedFrom(True, getChecked, setChecked, action)
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

        Public Function WithId(id As String) As CheckBoxBuilder Implements IID(Of CheckBoxBuilder).WithId
            _builder.WithId(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As CheckBoxBuilder Implements IID(Of CheckBoxBuilder).WithIdQ
            _builder.WithId([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As CheckBoxBuilder Implements IID(Of CheckBoxBuilder).WithIdMso
            _builder.WithId(mso)
            Return Me
        End Function

    End Class

End NameSpace