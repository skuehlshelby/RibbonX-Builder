Namespace RibbonAttributes
    Friend MustInherit Class RibbonAttribute
        Protected Const XmlTemplate As String = "{0}=""{1}"""
        Protected ReadOnly Name As AttributeName
        Protected ReadOnly CategoryMembers As ISet(Of AttributeName)

        Protected Sub New(name As AttributeName, ParamArray categoryMembers() As AttributeName)
            Me.Name = name
            Me.CategoryMembers = New HashSet(Of AttributeName)(categoryMembers) From {
                name
            }
        End Sub

        Public MustOverride ReadOnly Property Xml As String

        Public Overrides Function ToString() As String
            Return Xml
        End Function

        Public Shared Function ByCategory() As IEqualityComparer(Of RibbonAttribute)
            Return New CategoryComparer()
        End Function

        Public Shared Function ByCategoryMember(sampleCategoryMember As AttributeName) As IEquatable(Of RibbonAttribute)
            Return New NameComparer(sampleCategoryMember)
        End Function

        Private NotInheritable Class CategoryComparer
            Implements IEqualityComparer(Of RibbonAttribute)

            Public Overloads Function Equals(x As RibbonAttribute, y As RibbonAttribute) As Boolean Implements IEqualityComparer(Of RibbonAttribute).Equals
                Return x IsNot Nothing AndAlso y IsNot Nothing AndAlso x.CategoryMembers.Overlaps(y.CategoryMembers)
            End Function

            Public Overloads Function GetHashCode(obj As RibbonAttribute) As Integer Implements IEqualityComparer(Of RibbonAttribute).GetHashCode
                Dim hash As Integer

                Return obj.CategoryMembers.Aggregate(hash, Function(current, member) (current Xor member.GetHashCode()))
            End Function
        End Class

        Private NotInheritable Class NameComparer
            Implements IEquatable(Of RibbonAttribute)

            Private ReadOnly _name As AttributeName

            Public Sub New(name As AttributeName)
                _name = name
            End Sub

            Public Overloads Function Equals(other As RibbonAttribute) As Boolean Implements IEquatable(Of RibbonAttribute).Equals
                Return other IsNot Nothing AndAlso other.CategoryMembers.Contains(_name)
            End Function
        End Class

    End Class

End NameSpace