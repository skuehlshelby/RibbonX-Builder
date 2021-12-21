Namespace RibbonAttributes

    Friend MustInherit Class RibbonAttribute
        Implements IEquatable(Of RibbonAttribute)
        Protected Const XmlTemplate As String = "{0}=""{1}"""
        Protected ReadOnly Name As AttributeName
        Protected ReadOnly Category As AttributeCategory

        Protected Sub New(name As AttributeName, category As AttributeCategory)
            Me.Name = name
            Me.Category = category

            If Not Me.Category.Contains(name) Then
                Throw New ArgumentException($"Attribute {name} is not a member of {category}.")
            End If
        End Sub

        Public Function IsNamed(otherName As AttributeName) As Boolean
            Return Name.Equals(otherName)
        End Function

        Public Function IsExclusiveWith(other As AttributeName) As Boolean
            Return Category.Contains(other)
        End Function

        Public Function IsExclusiveWith(other As AttributeCategory) As Boolean
            Return Category.Equals(other)
        End Function

        Public Function IsExclusiveWith(other As RibbonAttribute) As Boolean
            Return Category.Equals(other.Category)
        End Function

        Public MustOverride ReadOnly Property Xml As String

        Public Overrides Function ToString() As String
            Return Xml
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Name.GetHashCode() Xor Category.GetHashCode()
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, RibbonAttribute))
        End Function

        Public Overloads Function Equals(other As RibbonAttribute) As Boolean Implements IEquatable(Of RibbonAttribute).Equals
            Return other IsNot Nothing AndAlso Name.Equals(other.Name) AndAlso Category.Equals(other.Category)
        End Function

        Public Shared Function CompareByCategory() As IEqualityComparer(Of RibbonAttribute)
            Return New AttributesAreEqualWhenTheyAreMutuallyExclusive()
        End Function

        Private NotInheritable Class AttributesAreEqualWhenTheyAreMutuallyExclusive
            Inherits EqualityComparer(Of RibbonAttribute)

            Public Overrides Function Equals(x As RibbonAttribute, y As RibbonAttribute) As Boolean
                Return x.IsExclusiveWith(y)
            End Function

            Public Overrides Function GetHashCode(obj As RibbonAttribute) As Integer
                Return obj.Category.GetHashCode()
            End Function
        End Class

    End Class

End NameSpace