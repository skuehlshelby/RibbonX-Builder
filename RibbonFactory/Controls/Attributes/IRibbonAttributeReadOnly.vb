Namespace Controls.Attributes

    Friend Interface IRibbonAttributeReadOnly(Of T)
        Inherits IRibbonAttribute

        Function GetValue() As T

    End Interface

End Namespace