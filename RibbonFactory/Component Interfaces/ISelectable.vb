Public Interface ISelectable
    Overloads Sub SelectItem(ByVal Index As Integer)
    Overloads Sub SelectItem(ByVal ID As String)
    ReadOnly Property SelectedItemIndex As Integer
End Interface
Public Interface ISelectable(Of Out T)
    Inherits ISelectable
    ReadOnly Property SelectedItem As T
End Interface
