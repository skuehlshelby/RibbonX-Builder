Imports RibbonX.InternalApi
Imports RibbonX.Properties

Friend MustInherit Class ContainerOfItems
    Inherits MutableContainer(Of IItem)
    Implements IItemTemplateable

    Protected ReadOnly Templates As ICollection(Of IItemTemplate) = New LinkedList(Of IItemTemplate)()

    Public Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
        MyBase.New(attributes, tag)
    End Sub

    Public Sub AddTemplate(template As IItemTemplate) Implements IItemTemplateable.AddTemplate
        Templates.Add(template)
    End Sub

    Public Sub AddTemplatedItem(item As Object) Implements IItemTemplateable.AddTemplatedItem
        For Each template As IItemTemplate In Templates
            If template.Match(item) Then
                Items.Add(template.Apply(item))
                Return
            End If
        Next

        Throw New Exception($"Failed to find a template for item {item}.")
    End Sub

    Public Sub RemoveTemplatedItem(item As Object) Implements IItemTemplateable.RemoveTemplatedItem
        For Each template As IItemTemplate In Templates
            If template.Match(item) Then
                Items.Remove(template.Apply(item))
                Return
            End If
        Next

        Throw New Exception($"Failed to find a template for item {item}.")
    End Sub

End Class
