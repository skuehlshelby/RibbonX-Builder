''' <summary>
''' Provides a mechanism to create <see cref="IItem"/>s from <see langword="Object"/>s. <br/>
''' Note: Returned <see cref="IItem"/>s are tracked by their <see cref="IRibbonElement.Id"/>,  <br/>
''' so ensure that the same <see langword="Object"/> will always produce an <see cref="IItem"/> with the same <see cref="IRibbonElement.Id"/>.
''' </summary>
Public Interface IItemTemplate
    Function Match(obj As Object) As Boolean
    Function Apply(obj As Object) As IItem
End Interface