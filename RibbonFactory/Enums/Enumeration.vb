Imports System.Reflection

Namespace Enums

    Public MustInherit Class Enumeration
        Implements IEquatable(Of Enumeration)
        Implements IComparable(Of Enumeration)
        Implements ICloneable

        Protected ReadOnly value As Integer
        Protected ReadOnly name As String

        Protected Sub New(value As Integer, name As String)
            Me.value = value
            Me.name = name
        End Sub

        Public Shared Function Values(Of T As Enumeration)() As T()
            Return GetType(T).
                GetProperties(BindingFlags.Public Or BindingFlags.Static Or BindingFlags.DeclaredOnly).
                Select(Function(p) p.GetValue(Nothing)).
                OfType(Of T)().
                ToArray()
        End Function

        Public Shared Function Parse(Of T As Enumeration)(value As Integer) As T
            Return Values(Of T)().Single(Function(v) v.value = value)
        End Function

        Public Shared Function Parse(Of T As Enumeration)(name As String) As T
            Return Values(Of T)().Single(Function(v) String.Equals(v.name, name, StringComparison.OrdinalIgnoreCase))
        End Function

        Public MustOverride Function Clone() As Object Implements ICloneable.Clone

#Region "Overrides And Comparison"

        Public Overrides Function ToString() As String
            Return name
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return value
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, Enumeration))
        End Function

        Public Overloads Function Equals(other As Enumeration) As Boolean Implements IEquatable(Of Enumeration).Equals
            Return other IsNot Nothing AndAlso other.value = value
        End Function

        Public Function CompareTo(other As Enumeration) As Integer Implements IComparable(Of Enumeration).CompareTo
            Return value.CompareTo(other.value)
        End Function

#End Region

#Region "Operators"

        Public Shared Operator =(left As Enumeration, right As Enumeration) As Boolean
            If left Is Nothing AndAlso right Is Nothing Then
                Return True
            ElseIf left Is Nothing OrElse right Is Nothing Then
                Return False
            Else
                Return left.Equals(right)
            End If
        End Operator

        Public Shared Operator <>(left As Enumeration, right As Enumeration) As Boolean
            Return Not left = right
        End Operator

        Public Shared Operator >(left As Enumeration, right As Enumeration) As Boolean
            Return left.CompareTo(right) > 0
        End Operator

        Public Shared Operator <(left As Enumeration, right As Enumeration) As Boolean
            Return left.CompareTo(right) < 0
        End Operator

        Public Shared Operator >=(left As Enumeration, right As Enumeration) As Boolean
            Return left.CompareTo(right) >= 0
        End Operator

        Public Shared Operator <=(left As Enumeration, right As Enumeration) As Boolean
            Return left.CompareTo(right) <= 0
        End Operator

#End Region

    End Class

End Namespace