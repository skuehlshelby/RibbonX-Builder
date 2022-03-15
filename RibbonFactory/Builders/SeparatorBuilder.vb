Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes

Namespace Builders

	Public NotInheritable Class SeparatorBuilder
		Implements IBuilder(Of Separator)
		Implements IID(Of SeparatorBuilder)
		Implements IInsert(Of SeparatorBuilder)
		Implements IVisible(Of SeparatorBuilder)

		Private ReadOnly _builder As ControlBuilder

		Public Sub New()
			Dim defaultProvider As IDefaultProvider = New Defaults(Of Separator)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
			attributeGroupBuilder.SetDefaults(defaultProvider)
			_builder = New ControlBuilder(attributeGroupBuilder)
		End Sub

		Public Sub New(template As Separator)
			Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
			attributeGroupBuilder.SetDefaults(template)
			attributeGroupBuilder.AddID(IdManager.GetID(Of Separator)())
			_builder = New ControlBuilder(attributeGroupBuilder)
		End Sub

		Public Sub New(template As RibbonElement)
			Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

			If defaultProvider IsNot Nothing Then
				Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
				Dim separatorAttributes As AttributeSet = New Defaults(Of Separator)().GetDefaults()
				Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) separatorAttributes.Contains(a)))
				Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
				attributeGroupBuilder.AddID(IdManager.GetID(Of Separator)())
				_builder = New ControlBuilder(attributeGroupBuilder)
			Else
				Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(Separator)}'")
			End If
		End Sub

		Public Function Build(Optional tag As Object = Nothing) As Separator Implements IBuilder(Of Separator).Build
			Return New Separator(_builder.Build(), tag)
		End Function

		Public Function WithId(id As String) As SeparatorBuilder Implements IID(Of SeparatorBuilder).WithId
			_builder.WithId(id)
			Return Me
		End Function

		Public Function WithIdQ([namespace] As String, id As String) As SeparatorBuilder Implements IID(Of SeparatorBuilder).WithIdQ
			_builder.WithId([namespace], id)
			Return Me
		End Function

		Public Function WithIdMso(mso As MSO) As SeparatorBuilder Implements IID(Of SeparatorBuilder).WithIdMso
			_builder.WithId(mso)
			Return Me
		End Function

		Public Function InsertBeforeMSO(builtInControl As MSO) As SeparatorBuilder Implements IInsert(Of SeparatorBuilder).InsertBefore
			_builder.InsertBefore(builtInControl)
			Return Me
		End Function

		Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As SeparatorBuilder Implements IInsert(Of SeparatorBuilder).InsertBefore
			_builder.InsertBefore(qualifiedControl)
			Return Me
		End Function

		Public Function InsertAfterMSO(builtInControl As MSO) As SeparatorBuilder Implements IInsert(Of SeparatorBuilder).InsertAfter
			_builder.InsertAfter(builtInControl)
			Return Me
		End Function

		Public Function InsertAfterQ(qualifiedControl As RibbonElement) As SeparatorBuilder Implements IInsert(Of SeparatorBuilder).InsertAfter
			_builder.InsertAfter(qualifiedControl)
			Return Me
		End Function

		Public Function Visible() As SeparatorBuilder Implements IVisible(Of SeparatorBuilder).Visible
			_builder.Visible()
			Return Me
		End Function

		Public Function Visible(callback As FromControl(Of Boolean)) As SeparatorBuilder Implements IVisible(Of SeparatorBuilder).Visible
			_builder.Visible(callback)
			Return Me
		End Function

		Public Function Invisible() As SeparatorBuilder Implements IVisible(Of SeparatorBuilder).Invisible
			_builder.Invisible()
			Return Me
		End Function

		Public Function Invisible(callback As FromControl(Of Boolean)) As SeparatorBuilder Implements IVisible(Of SeparatorBuilder).Invisible
			_builder.Invisible(callback)
			Return Me
		End Function

	End Class

End Namespace

