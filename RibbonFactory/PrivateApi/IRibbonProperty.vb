Imports RibbonX.Controls.Attributes

Namespace PrivateApi

    Friend Interface IRibbonProperty
        Inherits IEquatable(Of IRibbonProperty)

        ReadOnly Property Name As AttributeName
        ReadOnly Property Category As AttributeCategory
        ReadOnly Property IsReadOnly As Boolean
        ReadOnly Property IsDefault As Boolean
        Function GetBoxedValue() As Object
    End Interface

    Friend Interface IRibbonProperty(Of T)
        Inherits IRibbonProperty

        Function GetValue() As T
        Function SetValue(value As T) As Boolean
    End Interface

End Namespace