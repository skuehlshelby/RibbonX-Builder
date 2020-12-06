Public Interface IOrientable
    ReadOnly Property Orientation As BoxStyle
End Interface
Public Interface IOrientable(Of Out T)
    Function Horizontal() As T
    Function Vertical() As T
End Interface
