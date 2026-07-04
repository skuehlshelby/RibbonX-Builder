Namespace Images.BuiltIn

    Public Class ThreeDimensional
        Inherits ImageMSO

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property _3DBevelOptionsDialog As ThreeDimensional = New ThreeDimensional(1, NameOf(_3DBevelOptionsDialog))
        Public Shared ReadOnly Property _3DBevelPictureTopGallery As ThreeDimensional = New ThreeDimensional(2, NameOf(_3DBevelPictureTopGallery))
        Public Shared ReadOnly Property _3DDirectionGalleryClassic As ThreeDimensional = New ThreeDimensional(3, NameOf(_3DDirectionGalleryClassic))
        Public Shared ReadOnly Property _3DEffectColorPickerClassic As ThreeDimensional = New ThreeDimensional(4, NameOf(_3DEffectColorPickerClassic))
        Public Shared ReadOnly Property _3DEffectsGalleryClassic As ThreeDimensional = New ThreeDimensional(5, NameOf(_3DEffectsGalleryClassic))
        Public Shared ReadOnly Property _3DEffectsOnOffClassic As ThreeDimensional = New ThreeDimensional(6, NameOf(_3DEffectsOnOffClassic))
        Public Shared ReadOnly Property _3DExtrusionDepth144Classic As ThreeDimensional = New ThreeDimensional(7, NameOf(_3DExtrusionDepth144Classic))
        Public Shared ReadOnly Property _3DExtrusionDepth288Classic As ThreeDimensional = New ThreeDimensional(8, NameOf(_3DExtrusionDepth288Classic))
        Public Shared ReadOnly Property _3DExtrusionDepth36Classic As ThreeDimensional = New ThreeDimensional(9, NameOf(_3DExtrusionDepth36Classic))
        Public Shared ReadOnly Property _3DExtrusionDepth72Classic As ThreeDimensional = New ThreeDimensional(10, NameOf(_3DExtrusionDepth72Classic))
        Public Shared ReadOnly Property _3DExtrusionDepthGalleryClassic As ThreeDimensional = New ThreeDimensional(11, NameOf(_3DExtrusionDepthGalleryClassic))
        Public Shared ReadOnly Property _3DExtrusionDepthInfinityClassic As ThreeDimensional = New ThreeDimensional(12, NameOf(_3DExtrusionDepthInfinityClassic))
        Public Shared ReadOnly Property _3DExtrusionDepthNoneClassic As ThreeDimensional = New ThreeDimensional(13, NameOf(_3DExtrusionDepthNoneClassic))
        Public Shared ReadOnly Property _3DExtrusionDirectionClassic As ThreeDimensional = New ThreeDimensional(14, NameOf(_3DExtrusionDirectionClassic))
        Public Shared ReadOnly Property _3DExtrusionParallelClassic As ThreeDimensional = New ThreeDimensional(15, NameOf(_3DExtrusionParallelClassic))
        Public Shared ReadOnly Property _3DExtrusionPerspectiveClassic As ThreeDimensional = New ThreeDimensional(16, NameOf(_3DExtrusionPerspectiveClassic))
        Public Shared ReadOnly Property _3DLightingClassic As ThreeDimensional = New ThreeDimensional(17, NameOf(_3DLightingClassic))
        Public Shared ReadOnly Property _3DLightingDimClassic As ThreeDimensional = New ThreeDimensional(18, NameOf(_3DLightingDimClassic))
        Public Shared ReadOnly Property _3DLightingFlatClassic As ThreeDimensional = New ThreeDimensional(19, NameOf(_3DLightingFlatClassic))
        Public Shared ReadOnly Property _3DLightingGalleryClassic As ThreeDimensional = New ThreeDimensional(20, NameOf(_3DLightingGalleryClassic))
        Public Shared ReadOnly Property _3DLightingNormalClassic As ThreeDimensional = New ThreeDimensional(21, NameOf(_3DLightingNormalClassic))
        Public Shared ReadOnly Property _3DMaterialMetal As ThreeDimensional = New ThreeDimensional(22, NameOf(_3DMaterialMetal))
        Public Shared ReadOnly Property _3DMaterialMixed As ThreeDimensional = New ThreeDimensional(23, NameOf(_3DMaterialMixed))
        Public Shared ReadOnly Property _3DMaterialPlastic As ThreeDimensional = New ThreeDimensional(24, NameOf(_3DMaterialPlastic))
        Public Shared ReadOnly Property _3DPerspectiveDecrease As ThreeDimensional = New ThreeDimensional(25, NameOf(_3DPerspectiveDecrease))
        Public Shared ReadOnly Property _3DPerspectiveIncrease As ThreeDimensional = New ThreeDimensional(26, NameOf(_3DPerspectiveIncrease))
        Public Shared ReadOnly Property _3DRotationGallery As ThreeDimensional = New ThreeDimensional(27, NameOf(_3DRotationGallery))
        Public Shared ReadOnly Property _3DRotationOptionsDialog As ThreeDimensional = New ThreeDimensional(28, NameOf(_3DRotationOptionsDialog))
        Public Shared ReadOnly Property _3DStyle As ThreeDimensional = New ThreeDimensional(29, NameOf(_3DStyle))
        Public Shared ReadOnly Property _3DSurfaceMaterialClassic As ThreeDimensional = New ThreeDimensional(30, NameOf(_3DSurfaceMaterialClassic))
        Public Shared ReadOnly Property _3DSurfaceMaterialGalleryClassic As ThreeDimensional = New ThreeDimensional(31, NameOf(_3DSurfaceMaterialGalleryClassic))
        Public Shared ReadOnly Property _3DSurfaceMatteClassic As ThreeDimensional = New ThreeDimensional(32, NameOf(_3DSurfaceMatteClassic))
        Public Shared ReadOnly Property _3DSurfaceMetalClassic As ThreeDimensional = New ThreeDimensional(33, NameOf(_3DSurfaceMetalClassic))
        Public Shared ReadOnly Property _3DSurfacePlasticClassic As ThreeDimensional = New ThreeDimensional(34, NameOf(_3DSurfacePlasticClassic))
        Public Shared ReadOnly Property _3DSurfaceWireFrameClassic As ThreeDimensional = New ThreeDimensional(35, NameOf(_3DSurfaceWireFrameClassic))
        Public Shared ReadOnly Property _3DTiltDownClassic As ThreeDimensional = New ThreeDimensional(36, NameOf(_3DTiltDownClassic))
        Public Shared ReadOnly Property _3DTiltLeftClassic As ThreeDimensional = New ThreeDimensional(37, NameOf(_3DTiltLeftClassic))
        Public Shared ReadOnly Property _3DTiltRightClassic As ThreeDimensional = New ThreeDimensional(38, NameOf(_3DTiltRightClassic))
        Public Shared ReadOnly Property _3DTiltUpClassic As ThreeDimensional = New ThreeDimensional(39, NameOf(_3DTiltUpClassic))

        Public Overrides Function Clone() As Object
            Return New ThreeDimensional(value, name)
        End Function
    End Class

End Namespace