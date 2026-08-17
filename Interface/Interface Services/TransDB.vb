Imports Abovo.FileManager

Namespace Abovo

    Public Class TransDBManager

        Public ModelID As Integer

        Public Sub New(ByVal SetModelID As Integer)

            ModelID = SetModelID

        End Sub

        'Compatibility entry point retained because DataInterfaceTemplate,
        'NREditorInterface and older interface code already call this method.
        '
        'The actual logic now lives in the per-model TransactionalDBSynchroniser,
        'which mirrors the Summit_Compatibility VBA rules and also preserves the
        'older TransCopy_<RangeName> convention as a fallback.
        Public Shared Sub CheckTransDBActions(ByVal ModelID As Integer, ByVal TargRng As String)

            If ExcelModels Is Nothing Then Return
            If ModelID < 0 OrElse ModelID >= ExcelModels.Length Then Return
            If ExcelModels(ModelID) Is Nothing Then Return
            If ExcelModels(ModelID).TransDBSync Is Nothing Then Return

            ExcelModels(ModelID).TransDBSync.SynchroniseForNamedRange(TargRng)

        End Sub

    End Class

End Namespace
