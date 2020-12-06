Imports System.Reflection
'Translated to VB.NET from LosTechies.com -> Enumeration Classes
Public MustInherit Class Enumeration
    Implements IComparable

    Private Structure TThis
        Public Value As Integer
        Public DisplayName As String
    End Structure

    Private This As TThis
    Protected Sub New(ByVal Value As Integer, ByVal DisplayName As String)
        This = New TThis With {
                .Value = Value,
                .DisplayName = DisplayName
            }
    End Sub

    Public ReadOnly Property Value As Integer
        Get
            Return This.Value
        End Get
    End Property
    Public ReadOnly Property DisplayName As String
        Get
            Return This.DisplayName
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return This.DisplayName.Replace(" "c, Nothing)
    End Function
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse TypeOf obj IsNot Enumeration Then
            Return False
        Else
            Return This.Value.Equals(CType(obj, Enumeration).Value)
        End If
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return This.Value.GetHashCode()
    End Function

    Public Shared Function GetAll(Of T As Enumeration)() As IEnumerable(Of T)
        Dim Fields As FieldInfo() = GetType(T).GetFields(BindingFlags.Public Or BindingFlags.Static Or BindingFlags.DeclaredOnly)

        Return Fields.Select(Function(f) CType(f.GetValue(Nothing), T)).Cast(Of T)
    End Function

    Public Shared Function AbsoluteDifference(ByVal FirstValue As Enumeration, ByVal SecondValue As Enumeration) As Integer
        Return Math.Abs(FirstValue.Value - SecondValue.Value)
    End Function

    Public Shared Function FromValue(Of T As Enumeration)(ByVal Value As Integer) As T
        Dim MatchingItem As T = Parse(Of T, Integer)(Value, "Value", Function(Item) Item.Value = Value)
        Return MatchingItem
    End Function

    Public Shared Function FromDisplayName(Of T As Enumeration)(ByVal DisplayName As String) As T
        Dim MatchingItem As T = Parse(Of T, String)(DisplayName, "Display Name", Function(Item) Item.DisplayName = DisplayName)
        Return MatchingItem
    End Function

    Public Shared Function Parse(Of T As Enumeration, K)(ByVal Value As K, ByVal Description As String, ByVal Predicate As Func(Of T, Boolean)) As T
        Dim MatchingItem As T = GetAll(Of T).FirstOrDefault(Predicate)

        If MatchingItem Is Nothing Then
            Throw New ApplicationException(String.Format("'{0}' is not a valid {1} in {2}", Value, Description, MatchingItem.GetType))
        End If

        Return MatchingItem
    End Function
    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
        Return This.Value.CompareTo(CType(obj, Enumeration).Value)
    End Function
End Class
