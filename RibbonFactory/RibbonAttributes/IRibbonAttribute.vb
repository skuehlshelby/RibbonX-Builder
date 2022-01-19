Namespace RibbonAttributes

	Friend Interface IRibbonAttribute
		Inherits IEquatable(Of IRibbonAttribute)

		Event ValueChanged()

		Function IsNamed(name As AttributeName) As Boolean

		Function IsExclusiveWith(name As AttributeName) As Boolean

		Function IsExclusiveWith(category As AttributeCategory) As Boolean

	End Interface

End Namespace