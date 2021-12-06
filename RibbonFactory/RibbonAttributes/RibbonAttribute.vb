Namespace RibbonAttributes

    Friend MustInherit Class RibbonAttribute
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

        Public MustOverride ReadOnly Property Xml As String

        Public Overrides Function ToString() As String
            Return Xml
        End Function

        Public Shared Function ByCategory() As IEqualityComparer(Of RibbonAttribute)
            Return New AttributeComparer()
        End Function

        Public Shared Function ByCategory(category As AttributeCategory) As IEquatable(Of RibbonAttribute)
            Return New CategoryComparer(category)
        End Function

        Public Shared Function ByCategory(sampleCategoryMember As AttributeName) As IEquatable(Of RibbonAttribute)
            Return New NameComparer(sampleCategoryMember)
        End Function

        Private NotInheritable Class AttributeComparer
            Implements IEqualityComparer(Of RibbonAttribute)

            Public Overloads Function Equals(x As RibbonAttribute, y As RibbonAttribute) As Boolean Implements IEqualityComparer(Of RibbonAttribute).Equals
                Return x IsNot Nothing AndAlso y IsNot Nothing AndAlso x.Category.Equals(y.Category)
            End Function

            Public Overloads Function GetHashCode(obj As RibbonAttribute) As Integer Implements IEqualityComparer(Of RibbonAttribute).GetHashCode
                Return obj.Category.GetHashCode()
            End Function
        End Class

        Private NotInheritable Class CategoryComparer
            Implements IEquatable(Of RibbonAttribute)

            Private ReadOnly _category As AttributeCategory

            Public Sub New(category As AttributeCategory)
                _category = category
            End Sub

            Public Overloads Function Equals(other As RibbonAttribute) As Boolean Implements IEquatable(Of RibbonAttribute).Equals
                Return other IsNot Nothing AndAlso _category.Equals(other.Category)
            End Function
        End Class

        Private NotInheritable Class NameComparer
            Implements IEquatable(Of RibbonAttribute)

            Private ReadOnly _name As AttributeName

            Public Sub New(name As AttributeName)
                _name = name
            End Sub

            Public Overloads Function Equals(other As RibbonAttribute) As Boolean Implements IEquatable(Of RibbonAttribute).Equals
                Return other IsNot Nothing AndAlso other.Category.Contains(_name)
            End Function
        End Class

    End Class

End NameSpace