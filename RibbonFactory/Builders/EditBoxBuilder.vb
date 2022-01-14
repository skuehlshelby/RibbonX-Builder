Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities.Validation
Imports stdole

Namespace Builders

    Public NotInheritable Class EditBoxBuilder
        Implements IBuilder(Of EditBox)
        Implements IID(Of EditBoxBuilder)
        Implements IEnabled(Of EditBoxBuilder)
        Implements IVisible(Of EditBoxBuilder)
        Implements IInsert(Of EditBoxBuilder)
        Implements ILabelScreenTipSuperTip(Of EditBoxBuilder)
        Implements IKeyTip(Of EditBoxBuilder)
        Implements IImage(Of EditBoxBuilder)
        Implements IShowLabel(Of EditBoxBuilder)
        Implements IMaxLength(Of EditBoxBuilder)
        Implements ISizeString(Of EditBoxBuilder)
        Implements IOnChange(Of EditBox, EditBoxBuilder)
        Implements IValidateText(Of EditBoxBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private ReadOnly _validationRules As ICollection(Of IValidate(Of String))

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of EditBox)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = New ControlBuilder(attributeGroupBuilder)

            _validationRules = New List(Of IValidate(Of String))
        End Sub

        Public Sub New(template As EditBox)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(template)
            attributeGroupBuilder.AddID(IdManager.GetID(Of EditBox)())
            _builder = New ControlBuilder(attributeGroupBuilder)
            _validationRules = New List(Of IValidate(Of String))
        End Sub

        Public Sub New(template As RibbonElement)
            Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

            If defaultProvider IsNot Nothing Then
                Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
                Dim editBoxAttributes As AttributeSet = New DefaultProvider(Of EditBox)().GetDefaults()
                Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) editBoxAttributes.Contains(a)))
                Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
                attributeGroupBuilder.AddID(IdManager.GetID(Of EditBox)())
                _builder = New ControlBuilder(attributeGroupBuilder)
                _validationRules = New List(Of IValidate(Of String))
            Else
                Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(EditBox)}'")
            End If
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As EditBox Implements IBuilder(Of EditBox).Build
            Return New EditBox(_builder.Build(), _validationRules.ToArray(), tag:=tag)
        End Function

        Public Function Enabled() As EditBoxBuilder Implements IEnabled(Of EditBoxBuilder).Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As EditBoxBuilder Implements IEnabled(Of EditBoxBuilder).Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As EditBoxBuilder Implements IEnabled(Of EditBoxBuilder).Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As EditBoxBuilder Implements IEnabled(Of EditBoxBuilder).Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As EditBoxBuilder Implements IVisible(Of EditBoxBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As EditBoxBuilder Implements IVisible(Of EditBoxBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As EditBoxBuilder Implements IVisible(Of EditBoxBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As EditBoxBuilder Implements IVisible(Of EditBoxBuilder).Invisible
            _builder.Invisible(callback)
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As EditBoxBuilder Implements ILabelScreenTipSuperTip(Of EditBoxBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As EditBoxBuilder Implements ILabelScreenTipSuperTip(Of EditBoxBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As EditBoxBuilder Implements ILabelScreenTipSuperTip(Of EditBoxBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As EditBoxBuilder Implements ILabelScreenTipSuperTip(Of EditBoxBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As EditBoxBuilder Implements ILabelScreenTipSuperTip(Of EditBoxBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As EditBoxBuilder Implements ILabelScreenTipSuperTip(Of EditBoxBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As EditBoxBuilder Implements IKeyTip(Of EditBoxBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As EditBoxBuilder Implements IKeyTip(Of EditBoxBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As EditBoxBuilder Implements IImage(Of EditBoxBuilder).WithImage
            _builder.WithImage(image)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As EditBoxBuilder Implements IImage(Of EditBoxBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As EditBoxBuilder Implements IImage(Of EditBoxBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As EditBoxBuilder Implements IImage(Of EditBoxBuilder).WithImage
            _builder.WithImage(imagePath)
            Return Me
        End Function

        Public Function ShowLabel() As EditBoxBuilder Implements IShowLabel(Of EditBoxBuilder).ShowLabel
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As EditBoxBuilder Implements IShowLabel(Of EditBoxBuilder).ShowLabel
            _builder.ShowLabel(callback)
            Return Me
        End Function

        Public Function HideLabel() As EditBoxBuilder Implements IShowLabel(Of EditBoxBuilder).HideLabel
            _builder.HideLabel()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As EditBoxBuilder Implements IShowLabel(Of EditBoxBuilder).HideLabel
            _builder.HideLabel(callback)
            Return Me
        End Function

        Public Function WithMaximumInputLength(maxLength As Byte) As EditBoxBuilder Implements IMaxLength(Of EditBoxBuilder).WithMaximumInputLength
            _builder.WithMaxLength(maxLength)
            Return Me
        End Function

        Public Function ThatDoes(action As Action(Of EditBox), callback As TextChanged) As EditBoxBuilder Implements IOnChange(Of EditBox, EditBoxBuilder).ThatDoes
            _builder.ThatDoes(action, callback)
            Return Me
        End Function

        Public Function WithText(text As String, outgoing As FromControl(Of String)) As EditBoxBuilder
            _builder.GetTextFrom(text, outgoing)
            Return Me
        End Function

        Public Function WithText(incoming As TextChanged, outgoing As FromControl(Of String)) As EditBoxBuilder
            _builder.GetTextFrom(String.Empty, outgoing)
            Return Me
        End Function

        Public Function WithText(text As String, incoming As TextChanged, outgoing As FromControl(Of String)) As EditBoxBuilder
            _builder.GetTextFrom(text, outgoing)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As EditBoxBuilder Implements IInsert(Of EditBoxBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As EditBoxBuilder Implements IInsert(Of EditBoxBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As EditBoxBuilder Implements IInsert(Of EditBoxBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As EditBoxBuilder Implements IInsert(Of EditBoxBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

        Public Function WithTextValidationRule(rule As Predicate(Of String)) As EditBoxBuilder Implements IValidateText(Of EditBoxBuilder).WithTextValidationRule
            Return WithTextValidationRule(rule, String.Empty)
        End Function

        Public Function WithTextValidationRule(rule As Predicate(Of String), failureMessage As String) As EditBoxBuilder Implements IValidateText(Of EditBoxBuilder).WithTextValidationRule
            _validationRules.Add(New TextValidator(rule, failureMessage))
            Return Me
        End Function

        Public Function AsWideAs(sizeString As String) As EditBoxBuilder Implements ISizeString(Of EditBoxBuilder).AsWideAs
            _builder.WithSize(sizeString)
            Return Me
        End Function

        Public Function WithId(id As String) As EditBoxBuilder Implements IID(Of EditBoxBuilder).WithId
            _builder.WithId(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As EditBoxBuilder Implements IID(Of EditBoxBuilder).WithIdQ
            _builder.WithId([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As EditBoxBuilder Implements IID(Of EditBoxBuilder).WithIdMso
            _builder.WithId(mso)
            Return Me
        End Function

    End Class

End Namespace