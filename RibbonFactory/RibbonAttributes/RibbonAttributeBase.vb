Imports RibbonFactory.RibbonAttributes.Categories.Enabled

Namespace RibbonAttributes
    Public MustInherit Class RibbonAttribute
        Implements IRibbonAttribute

        Protected Const XML_TEMPLATE As String = "{0}=""{1}"""

        Public MustOverride ReadOnly Property XML As String Implements IRibbonAttribute.XML

        Public Overrides Function ToString() As String
            Return XML
        End Function

        Public MustOverride Overrides Function GetHashCode() As Integer

        Public MustOverride Overrides Function Equals(obj As Object) As Boolean

        Protected Shared Function CamelCase(input As Object) As String
            Return input.ToString().Substring(0, 1).ToLower() & input.ToString().Substring(1)
        End Function

        Public Overloads Shared Operator =(left As RibbonAttribute, right As RibbonAttribute) As Boolean
            Return left.Equals(right)
        End Operator

        Public Overloads Shared Operator =(left As RibbonAttribute, right As Object) As Boolean
            Return left.Equals(right)
        End Operator

        Public Overloads Shared Operator =(left As Object, right As RibbonAttribute) As Boolean
            Return right.Equals(left)
        End Operator

        Public Overloads Shared Operator <>(left As RibbonAttribute, right As RibbonAttribute) As Boolean
            Return Not left.Equals(right)
        End Operator

        Public Overloads Shared Operator <>(left As RibbonAttribute, right As Object) As Boolean
            Return Not left.Equals(right)
        End Operator

        Public Overloads Shared Operator <>(left As Object, right As RibbonAttribute) As Boolean
            Return Not right.Equals(left)
        End Operator
    End Class
End Namespace