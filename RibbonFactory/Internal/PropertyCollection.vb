
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace InternalApi

    Friend NotInheritable Class PropertyCollection
        Implements IPropertyCollection

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private ReadOnly events As EventHandlerList = New EventHandlerList()
        Private ReadOnly properties As ISet(Of IRibbonProperty) = New HashSet(Of IRibbonProperty)()
        Private ReadOnly forwardingRules As IDictionary(Of IPropertyCategory, IPropertyCategory) = New Dictionary(Of IPropertyCategory, IPropertyCategory)()

        Public Sub New(ParamArray properties() As IRibbonProperty)
            Me.properties = New HashSet(Of IRibbonProperty)(properties)
        End Sub

        Public ReadOnly Property Count As Integer Implements ICollection(Of IRibbonProperty).Count
            Get
                Return properties.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of IRibbonProperty).IsReadOnly
            Get
                Return False
            End Get
        End Property

        Public Sub Add(item As IRibbonProperty) Implements ICollection(Of IRibbonProperty).Add
            properties.Remove(item)
            properties.Add(item)
        End Sub

        Public Sub Clear() Implements ICollection(Of IRibbonProperty).Clear
            properties.Clear()
        End Sub

        Public Function Contains(item As IRibbonProperty) As Boolean Implements ICollection(Of IRibbonProperty).Contains
            Return properties.Contains(item)
        End Function

        Public Sub CopyTo(array() As IRibbonProperty, arrayIndex As Integer) Implements ICollection(Of IRibbonProperty).CopyTo
            properties.CopyTo(array, arrayIndex)
        End Sub

        Public Function Remove(item As IRibbonProperty) As Boolean Implements ICollection(Of IRibbonProperty).Remove
            Return properties.Remove(item)
        End Function

        Public Sub Merge(other As IPropertyCollection) Implements IPropertyCollection.Merge
            For Each ribbonProperty As IRibbonProperty In other
                If Not Equals(ribbonProperty.Category, Category.IdType) Then
                    If Contains(ribbonProperty) Then
                        Remove(ribbonProperty)
                        Add(ribbonProperty)
                    End If
                End If
            Next
        End Sub

        Public Sub Forward(from As IPropertyCategory, [to] As IPropertyCategory) Implements IPropertyCollection.Forward
            forwardingRules.Add(from, [to])
        End Sub

        Private Function ResolveForwardingRule(category As IPropertyCategory) As IPropertyCategory
            Return If(forwardingRules.ContainsKey(category), forwardingRules(category), category)
        End Function

        Private Function TryFind(category As IPropertyCategory, <Out()> ByRef result As IRibbonProperty) As Boolean
            For Each ribbonProperty As IRibbonProperty In properties
                If Equals(ribbonProperty.Category, category) Then
                    result = ribbonProperty
                    Return True
                End If
            Next

            result = Nothing
            Return False
        End Function

        Public Function [Get](Of T)(category As IPropertyCategory(Of T)) As T Implements IPropertyCollection.Get
            Dim target As IRibbonProperty = Nothing

            If TryFind(ResolveForwardingRule(category), target) Then
                Return CType(target, IRibbonProperty(Of T)).GetValue()
            Else
                Throw NoResults(Of T)(category)
            End If
        End Function

        Public Sub [Set](Of T)(value As T, category As IPropertyCategory(Of T), <CallerMemberName> Optional callerMemberName As String = Nothing) Implements IPropertyCollection.Set
            Dim target As IRibbonProperty = Nothing

            If TryFind(ResolveForwardingRule(category), target) Then
                If target.IsReadOnly Then
                    Throw New PropertyNotSettableException(callerMemberName)
                Else
                    If CType(target, IRibbonProperty(Of T)).SetValue(value) Then
                        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(callerMemberName))
                    End If
                End If
            Else
                Throw NoResults(Of T)(category)
            End If
        End Sub

        Private Function NoResults(Of T)(category As IPropertyCategory) As Exception
            Return New Exception($"Failed to find any objects of type '{GetType(IRibbonProperty(Of T)).Name}' and category '{category}'.")
        End Function

        Public Function Raw(category As IPropertyCategory) As IRibbonProperty Implements IPropertyCollection.Raw
            Dim target As IRibbonProperty = Nothing

            If TryFind(category, target) Then
                Return target
            Else
                Throw New Exception($"Failed to find any objects of category '{category}'.")
            End If
        End Function

        Public Function Raw(ParamArray categories() As IPropertyCategory) As IRibbonProperty() Implements IPropertyCollection.Raw
            Return properties.Where(Function(p) categories.Contains(p.Category)).ToArray()
        End Function

        Public Sub [AddHandler](eventName As String, handler As [Delegate]) Implements IPropertyCollection.AddHandler
            events.AddHandler(eventName, handler)
        End Sub

        Public Sub [RemoveHandler](eventName As String, handler As [Delegate]) Implements IPropertyCollection.RemoveHandler
            events.RemoveHandler(eventName, handler)
        End Sub

        Public Sub [RaiseEvent](Of TEventArgs As EventArgs)(eventName As String, sender As Object, e As TEventArgs) Implements IPropertyCollection.RaiseEvent
            Dim handler As Object = events(eventName)

            If handler IsNot Nothing Then
                If TypeOf handler Is EventHandler Then
                    CType(handler, EventHandler).Invoke(sender, e)
                Else
                    CType(handler, EventHandler(Of TEventArgs)).Invoke(sender, e)
                End If
            End If
        End Sub

        Private Function Except(ParamArray categories() As IPropertyCategory) As IEnumerable(Of IRibbonProperty)
            Return properties.Where(Function(p) Not categories.Contains(p.Category))
        End Function

        Public Overloads Function Clone() As Object Implements ICloneable.Clone
            Return New PropertyCollection(properties.Select(Function(p) p.Clone()).OfType(Of IRibbonProperty)().ToArray())
        End Function

        Public Function GetEnumerator() As IEnumerator(Of IRibbonProperty) Implements IEnumerable(Of IRibbonProperty).GetEnumerator
            Return properties.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return properties.GetEnumerator()
        End Function

        Public Overrides Function ToString() As String
            Return $"{NameOf(PropertyCollection)} Count:={properties.Count})"
        End Function

        Public Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String Implements IPropertyCollection.ToXml
            Select Case exclude
                Case ExcludedAttributes.Size
                    Return String.Join(" ", Except(Category.Size).Select(Function([property]) $"{[property].ToXml()}"))
                Case ExcludedAttributes.SizeAndVisibility
                    Return String.Join(" ", Except(Category.Size, Category.Visibility).Select(Function([property]) $"{[property].ToXml()}"))
                Case Else
                    Return String.Join(" ", properties.Select(Function([property]) $"{[property].ToXml()}"))
            End Select
        End Function
    End Class

End Namespace