Imports RibbonX.InternalApi
Imports RibbonX.Utilities

Namespace Controls.Base

    Friend MustInherit Class ReadOnlyContainer(Of TRibbonElement As IRibbonElement)
        Inherits RibbonElement
        Implements IReadOnlyCollection(Of TRibbonElement)

        Protected ReadOnly Items As ICollection(Of TRibbonElement)

        Protected Sub New(attributes As IPropertyCollection, items As ICollection(Of TRibbonElement), Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
            Me.Items = If(items, Array.Empty(Of TRibbonElement)())

            For Each item As TRibbonElement In items
                If item IsNot Nothing Then
                    AddHandler item.PropertyChanged, AddressOf OnPropertyChanged
                    AddHandler item.RefreshNeeded, AddressOf OnRefreshNeeded
                End If
            Next
        End Sub

        Public ReadOnly Property Count As Integer Implements IReadOnlyCollection(Of TRibbonElement).Count
            Get
                Return Items.Count
            End Get
        End Property

        Public Overrides Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String
            Return $"<{[GetType]().Name.CamelCase()} {Attributes.ToXml(exclude)}>{String.Join(Environment.NewLine, Items.Select(Function(child) child.ToXml(exclude)))}</{[GetType].Name.CamelCase()}>"
        End Function

        Public Function GetEnumerator() As IEnumerator(Of TRibbonElement) Implements IEnumerable(Of TRibbonElement).GetEnumerator
            Return Items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return Items.GetEnumerator()
        End Function
    End Class

End Namespace