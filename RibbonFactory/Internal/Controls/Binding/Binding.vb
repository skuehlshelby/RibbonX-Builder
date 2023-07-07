Imports System.ComponentModel
Imports System.Linq.Expressions

Friend NotInheritable Class BindingFactory
    Public Shared ReadOnly Instance As BindingFactory = New BindingFactory()

    Private ReadOnly bindings As ICollection(Of IBinding) = New LinkedList(Of IBinding)()

    Private Sub New()
    End Sub

    Private Interface IBinding
        Function IsAlive() As Boolean
        Sub OnPropertyUpdate(sender As Object, e As PropertyChangedEventArgs)
    End Interface

    Private Class Binding(Of T As Class, K As Class)
        Implements IBinding
        Private ReadOnly leftRef As WeakReference(Of T)
        Private ReadOnly rightRef As WeakReference(Of K)
        Private ReadOnly propertyName As String
        Private ReadOnly action As Action(Of T, K)

        Public Sub New(left As T, right As K, propertyName As String, action As Action(Of T, K))
            If left Is Nothing Then
                Throw New ArgumentNullException(NameOf(left))
            End If

            If right Is Nothing Then
                Throw New ArgumentNullException(NameOf(right))
            End If

            If propertyName Is Nothing Then
                Throw New ArgumentNullException(NameOf(propertyName))
            End If

            If action Is Nothing Then
                Throw New ArgumentNullException(NameOf(action))
            End If

            leftRef = New WeakReference(Of T)(left)
            rightRef = New WeakReference(Of K)(right)
            Me.propertyName = propertyName
            Me.action = action
        End Sub

        Public Sub OnPropertyUpdate(sender As Object, e As PropertyChangedEventArgs) Implements IBinding.OnPropertyUpdate
            If Equals(e.PropertyName, propertyName) Then
                Dim left As T = Nothing
                Dim right As K = Nothing
                If leftRef.TryGetTarget(left) AndAlso rightRef.TryGetTarget(right) Then
                    action.Invoke(left, right)
                End If
            End If
        End Sub

        Public Function IsAlive() As Boolean Implements IBinding.IsAlive
            Dim left As T = Nothing
            Dim right As K = Nothing
            Return leftRef.TryGetTarget(left) AndAlso rightRef.TryGetTarget(right)
        End Function
    End Class

    Public Sub OnPropertyUpdate(sender As Object, e As PropertyChangedEventArgs)
        Dim dead As ICollection(Of IBinding) = bindings.Where(Function(b) Not b.IsAlive()).ToArray()

        For Each binding As IBinding In dead
            bindings.Remove(binding)
        Next

        For Each binding As IBinding In bindings
            binding.OnPropertyUpdate(sender, e)
        Next
    End Sub

    Public Sub Create(Of TElement As {IRibbonElement, Class}, TOther As Class)(element As TElement, other As TOther, expression As Expression(Of Action(Of TElement, TOther)))
        Dim operation As MethodCallExpression = DirectCast(expression.Body, MethodCallExpression)
        Dim left As ParameterExpression = DirectCast(operation.Object, ParameterExpression)
        Dim right As MemberExpression = operation.Arguments.OfType(Of MemberExpression).Single()
        Dim name As String = right.Member.Name

        Dim action As Action(Of TElement, TOther) = expression.Compile()
        action.Invoke(element, other)

        If left.Type.Equals(GetType(TElement)) Then
            If TypeOf other Is INotifyPropertyChanged Then
                bindings.Add(New Binding(Of TElement, TOther)(element, other, name, action))
                AddHandler DirectCast(other, INotifyPropertyChanged).PropertyChanged, AddressOf OnPropertyUpdate
            End If
        Else
            bindings.Add(New Binding(Of TElement, TOther)(element, other, name, action))
            AddHandler element.PropertyChanged, AddressOf OnPropertyUpdate
        End If
    End Sub

End Class