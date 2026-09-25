Option Strict On

Imports System.Runtime.CompilerServices

Namespace Abovo.WorkbookEngines
    'Opening boundary only. Import-source and existing legacy callers remain
    'ordinary DevExpress opens; the main model window activates this before UI
    'creation. A live model is never silently transferred to another engine.
    Friend NotInheritable Class ModelEngineSelection
        Private NotInheritable Class SelectionNote
            Public Text As String
        End Class
        Private Shared ReadOnly notes As New ConditionalWeakTable(Of FileManager.ExcelModel, SelectionNote)()

        Friend Shared Function Status(model As FileManager.ExcelModel) As String
            If model Is Nothing OrElse model.IsClosing Then Return "No open model"
            If model.ChangeManager IsNot Nothing AndAlso model.ChangeManager.HasEngineEditingTrial Then
                Return model.ChangeManager.CalculationOwnerDescription
            End If
            Dim note As SelectionNote = Nothing
            Return If(notes.TryGetValue(model, note), "DevExpress — " & note.Text, "DevExpress")
        End Function

        Private Shared Sub UseDevExpress(model As FileManager.ExcelModel, reason As String)
            notes.GetOrCreateValue(model).Text = reason
            SystemMessageManager.Publish(model.ModelID, "Calculation engine: DevExpress. " & reason,
                SystemMessageSeverity.Information, "Calculation engine", model.FileName)
        End Sub

        Friend Shared Sub Activate(model As FileManager.ExcelModel, options As WorkbookEngineOptions,
                                   Optional factory As Func(Of WorkbookEnginePreference, IWorkbookCalculationBackend) = Nothing)
            If model Is Nothing OrElse model.IsClosing OrElse model.WB Is Nothing OrElse model.ChangeManager Is Nothing Then
                Throw New InvalidOperationException("The model is not available for calculation-engine selection.")
            End If
            If model.ChangeManager.HasEngineEditingTrial Then Throw New InvalidOperationException("Reopen the model to change its calculation engine.")
            If options Is Nothing Then Throw New ArgumentNullException(NameOf(options))
            If options.Preference = WorkbookEnginePreference.DevExpressOnly Then
                UseDevExpress(model, "Selected in Options.")
                Return
            End If
            Dim reason As String = Nothing
            If Not String.IsNullOrWhiteSpace(model.RecoverySourcePath) Then
                reason = "Recovery copies use DevExpress so Save As can restore the original workbook format."
            ElseIf model.IsDirty OrElse model.ChangeManager.GetHistoryTable().Rows.Count > 0 Then
                reason = "This model already has changes; engine selection applies only to a clean opening."
            ElseIf Not {".xlsb", ".xlsm", ".xlsx"}.Contains(IO.Path.GetExtension(model.FileName), StringComparer.OrdinalIgnoreCase) Then
                reason = "This workbook filename format uses the existing DevExpress route."
            Else
                Dim stress = model.WB.DefinedNames.GetDefinedName("StressTestMode")?.Range
                If stress IsNot Nothing AndAlso String.Equals(stress(0, 0).Value.TextValue, "Y", StringComparison.OrdinalIgnoreCase) Then
                    reason = "Stress Test mode uses DevExpress in this stage of engine integration."
                End If
            End If
            If reason IsNot Nothing Then
                If options.Preference = WorkbookEnginePreference.ExcelRequired Then Throw New NotSupportedException(reason)
                UseDevExpress(model, reason)
                Return
            End If

            Dim session As WorkbookCalculationSession = Nothing
            Try
                'The existing session owns capability/security checks and safe
                'opening-only fallback. Failure/timeout during cleanup propagates.
                session = WorkbookCalculationSession.OpenAsync(model.FileName, options, factory:=factory).GetAwaiter().GetResult()
                If session.EngineName = "DevExpress" Then
                    Dim fallback = session.FallbackReason
                    session.CloseAsync().GetAwaiter().GetResult()
                    session = Nothing
                    UseDevExpress(model, "Compatible Excel was unavailable. " & fallback)
                    Return
                End If
                Dim first = If(model.WB.Worksheets.Contains("Check Sheet"), "Check Sheet", model.WB.Worksheets(0).Name)
                Dim result = session.CalculateAndReadAsync(session.Revision, WorkbookCalculationKind.Rebuild,
                    {New WorkbookReadArea(first, 0, 0, 1, 1)}).GetAwaiter().GetResult()
                model.ChangeManager.BindEngineEditingTrial(session, result)
                session = Nothing 'Only the model may close the adopted owner.
                model.ChangeManager.AcceptInitialEngineCalculation()
                SystemMessageManager.Publish(model.ModelID,
                    "Calculation engine: " & Status(model) & ". Structural edits, imports and Stress Test still require reopening with DevExpress. Save As currently keeps the same file format and requires a new filename.",
                    SystemMessageSeverity.Information, "Calculation engine", model.FileName)
            Catch openingError As Exception
                If session IsNot Nothing Then
                    Try
                        session.CloseAsync().GetAwaiter().GetResult()
                    Catch cleanupError As Exception
                        Throw New AggregateException("Calculation-engine opening failed and its cleanup also reported a problem.",
                            openingError, cleanupError)
                    End Try
                End If
                Throw
            End Try
        End Sub
    End Class
End Namespace
