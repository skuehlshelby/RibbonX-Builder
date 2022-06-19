Namespace Controls.Attributes

    Friend Interface IRibbonAttributeReadWrite(Of T)
        Inherits IRibbonAttributeReadOnly(Of T)

        Sub SetValue(value As T)

    End Interface

End Namespace