Namespace Component_Interfaces
    Public Interface ISelectable
        Overloads Sub SelectItem(Index As Integer)
        Overloads Sub SelectItem(ID As String)
        ReadOnly Property SelectedItemIndex As Integer
    End Interface
    Public Interface ISelectable(Of Out T)
        Inherits ISelectable
        ReadOnly Property SelectedItem As T
    End Interface
End NameSpace