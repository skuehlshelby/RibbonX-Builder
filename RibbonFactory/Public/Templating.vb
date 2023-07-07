


Public Interface IItemTemplate
    Function Match(obj As Object) As Boolean
    Function Apply(obj As Object) As IItem
End Interface