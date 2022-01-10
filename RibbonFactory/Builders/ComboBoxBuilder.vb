Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders

	Public NotInheritable Class ComboBoxBuilder
		Implements IBuilder(Of ComboBox)
		Implements IID(Of ComboBoxBuilder)
		Implements IEnabled(Of ComboBoxBuilder)
		Implements IVisible(Of ComboBoxBuilder)
		Implements ILabelScreenTipSuperTip(Of ComboBoxBuilder)
		Implements IShowLabel(Of ComboBoxBuilder)
		Implements IKeyTip(Of ComboBoxBuilder)
		Implements IImage(Of ComboBoxBuilder)
		Implements IShowImage(Of ComboBoxBuilder)
		Implements IMaxLength(Of ComboBoxBuilder)
		Implements IInsert(Of ComboBoxBuilder)

		Private ReadOnly _builder As ControlBuilder

		Public Sub New()
			Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of ComboBox)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(defaultProvider.GetDefaults())
			_builder = New ControlBuilder(attributeGroupBuilder)
		End Sub

		Public Sub New(template As ComboBox)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
			attributeGroupBuilder.SetDefaults(template)
			attributeGroupBuilder.AddID(IdManager.GetID(Of ComboBox)())
			_builder = New ControlBuilder(attributeGroupBuilder)
		End Sub

		Public Sub New(template As Object)
			Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

			If defaultProvider IsNot Nothing Then
				Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
				Dim comboBoxAttributes As AttributeSet = New DefaultProvider(Of ComboBox)().GetDefaults()
				Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) comboBoxAttributes.Contains(a)))
				Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
				attributeGroupBuilder.AddID(IdManager.GetID(Of ComboBox)())
				_builder = New ControlBuilder(attributeGroupBuilder)
			Else
				Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(ComboBox)}'")
			End If
		End Sub

		Public Function Build(Optional tag As Object = Nothing) As ComboBox Implements IBuilder(Of ComboBox).Build
			Throw New NotImplementedException()
		End Function

		Public Function Enabled() As ComboBoxBuilder Implements IEnabled(Of ComboBoxBuilder).Enabled
			_builder.Enabled()
			Return Me
		End Function

		Public Function Enabled(callback As FromControl(Of Boolean)) As ComboBoxBuilder Implements IEnabled(Of ComboBoxBuilder).Enabled
			_builder.Enabled(callback)
			Return Me
		End Function

		Public Function Disabled() As ComboBoxBuilder Implements IEnabled(Of ComboBoxBuilder).Disabled
			_builder.Disabled()
			Return Me
		End Function

		Public Function Disabled(callback As FromControl(Of Boolean)) As ComboBoxBuilder Implements IEnabled(Of ComboBoxBuilder).Disabled
			_builder.Disabled(callback)
			Return Me
		End Function

		Public Function Visible() As ComboBoxBuilder Implements IVisible(Of ComboBoxBuilder).Visible
			_builder.Visible()
			Return Me
		End Function

		Public Function Visible(callback As FromControl(Of Boolean)) As ComboBoxBuilder Implements IVisible(Of ComboBoxBuilder).Visible
			_builder.Visible(callback)
			Return Me
		End Function

		Public Function Invisible() As ComboBoxBuilder Implements IVisible(Of ComboBoxBuilder).Invisible
			_builder.Invisible()
			Return Me
		End Function

		Public Function Invisible(callback As FromControl(Of Boolean)) As ComboBoxBuilder Implements IVisible(Of ComboBoxBuilder).Invisible
			_builder.Invisible(callback)
			Return Me
		End Function

		Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ComboBoxBuilder Implements ILabelScreenTipSuperTip(Of ComboBoxBuilder).WithLabel
			_builder.WithLabel(label, copyToScreenTip)
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ComboBoxBuilder Implements ILabelScreenTipSuperTip(Of ComboBoxBuilder).WithLabel
			_builder.WithLabel(label, callback, copyToScreenTip)
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function WithScreenTip(screenTip As String) As ComboBoxBuilder Implements ILabelScreenTipSuperTip(Of ComboBoxBuilder).WithScreenTip
			_builder.WithScreentip(screenTip)
			Return Me
		End Function

		Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ComboBoxBuilder Implements ILabelScreenTipSuperTip(Of ComboBoxBuilder).WithScreenTip
			_builder.WithScreentip(screenTip, callback)
			Return Me
		End Function

		Public Function WithSuperTip(superTip As String) As ComboBoxBuilder Implements ILabelScreenTipSuperTip(Of ComboBoxBuilder).WithSuperTip
			_builder.WithSupertip(superTip)
			Return Me
		End Function

		Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ComboBoxBuilder Implements ILabelScreenTipSuperTip(Of ComboBoxBuilder).WithSuperTip
			_builder.WithSupertip(superTip, callback)
			Return Me
		End Function

		Public Function ShowLabel() As ComboBoxBuilder Implements IShowLabel(Of ComboBoxBuilder).ShowLabel
			_builder.ShowLabel()
			Return Me
		End Function

		Public Function ShowLabel(callback As FromControl(Of Boolean)) As ComboBoxBuilder Implements IShowLabel(Of ComboBoxBuilder).ShowLabel
			_builder.ShowLabel(callback)
			Return Me
		End Function

		Public Function HideLabel() As ComboBoxBuilder Implements IShowLabel(Of ComboBoxBuilder).HideLabel
			_builder.HideLabel()
			Return Me
		End Function

		Public Function HideLabel(callback As FromControl(Of Boolean)) As ComboBoxBuilder Implements IShowLabel(Of ComboBoxBuilder).HideLabel
			_builder.HideLabel(callback)
			Return Me
		End Function

		Public Function WithKeyTip(keyTip As KeyTip) As ComboBoxBuilder Implements IKeyTip(Of ComboBoxBuilder).WithKeyTip
			_builder.WithKeyTip(keyTip)
			Return Me
		End Function

		Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As ComboBoxBuilder Implements IKeyTip(Of ComboBoxBuilder).WithKeyTip
			_builder.WithKeyTip(keyTip, callback)
			Return Me
		End Function

		Public Function WithImage(image As ImageMSO) As ComboBoxBuilder Implements IImage(Of ComboBoxBuilder).WithImage
			_builder.WithImage(image)
			_builder.ShowImage()
			Return Me
		End Function

		Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As ComboBoxBuilder Implements IImage(Of ComboBoxBuilder).WithImage
			_builder.WithImage(image, callback)
			_builder.ShowImage()
			Return Me
		End Function

		Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As ComboBoxBuilder Implements IImage(Of ComboBoxBuilder).WithImage
			_builder.WithImage(image, callback)
			Return Me
		End Function

		Public Function WithImage(imagePath As String) As ComboBoxBuilder Implements IImage(Of ComboBoxBuilder).WithImage
			_builder.WithImage(imagePath)
			_builder.ShowImage()
			Return Me
		End Function

		Public Function ShowImage() As ComboBoxBuilder Implements IShowImage(Of ComboBoxBuilder).ShowImage
			_builder.ShowImage()
			Return Me
		End Function

		Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As ComboBoxBuilder Implements IShowImage(Of ComboBoxBuilder).ShowImage
			_builder.ShowImage(getShowImage)
			Return Me
		End Function

		Public Function HideImage() As ComboBoxBuilder Implements IShowImage(Of ComboBoxBuilder).HideImage
			_builder.HideImage()
			Return Me
		End Function

		Public Function HideImage(getShowImage As FromControl(Of Boolean)) As ComboBoxBuilder Implements IShowImage(Of ComboBoxBuilder).HideImage
			_builder.HideImage(getShowImage)
			Return Me
		End Function

		Public Function WithMaximumInputLength(maxLength As Byte) As ComboBoxBuilder Implements IMaxLength(Of ComboBoxBuilder).WithMaximumInputLength
			_builder.WithMaxLength(maxLength)
			Return Me
		End Function

		Public Function InsertBeforeMSO(builtInControl As MSO) As ComboBoxBuilder Implements IInsert(Of ComboBoxBuilder).InsertBefore
			_builder.InsertBefore(builtInControl)
			Return Me
		End Function

		Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ComboBoxBuilder Implements IInsert(Of ComboBoxBuilder).InsertBefore
			_builder.InsertBefore(qualifiedControl)
			Return Me
		End Function

		Public Function InsertAfterMSO(builtInControl As MSO) As ComboBoxBuilder Implements IInsert(Of ComboBoxBuilder).InsertAfter
			_builder.InsertAfter(builtInControl)
			Return Me
		End Function

		Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ComboBoxBuilder Implements IInsert(Of ComboBoxBuilder).InsertAfter
			_builder.InsertAfter(qualifiedControl)
			Return Me
		End Function

		Public Function WithId(id As String) As ComboBoxBuilder Implements IID(Of ComboBoxBuilder).WithId
			_builder.WithId(id)
			Return Me
		End Function

		Public Function WithIdQ([namespace] As String, id As String) As ComboBoxBuilder Implements IID(Of ComboBoxBuilder).WithIdQ
			_builder.WithId([namespace], id)
			Return Me
		End Function

		Public Function WithIdMso(mso As MSO) As ComboBoxBuilder Implements IID(Of ComboBoxBuilder).WithIdMso
			_builder.WithId(mso)
			Return Me
		End Function

	End Class

End Namespace