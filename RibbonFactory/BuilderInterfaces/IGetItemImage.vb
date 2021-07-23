Imports stdole

Namespace BuilderInterfaces

    Public Interface IGetItemImage(Of T)

        Function GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp)) As T

    End Interface

End Namespace