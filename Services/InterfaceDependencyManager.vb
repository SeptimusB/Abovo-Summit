Imports System.Collections.Generic
Imports System.Linq
Imports DevExpress.Spreadsheet

Namespace Abovo

    ''' <summary>
    ''' Tracks which DataInterfaceTemplate sections depend upon which workbook worksheets.
    '''
    ''' The registry deliberately stores WeakReference objects rather than live controls or
    ''' UnboundSource instances.  This prevents hidden/closed interfaces being kept alive
    ''' solely because they were once registered as worksheet consumers.
    ''' </summary>
    Public Class InterfaceDependencyManager

        Private ReadOnly ModelID As Integer
        Private ReadOnly SyncRoot As New Object
        Private ReadOnly WorksheetSections As New Dictionary(Of String, List(Of SectionDependency))(StringComparer.OrdinalIgnoreCase)

        Private Class SectionDependency
            Public Owner As WeakReference
            Public SectionID As Integer
        End Class

        Public Sub New(ByVal SetModelID As Integer)
            ModelID = SetModelID
        End Sub

        Public Sub RegisterSection(ByVal Owner As DataInterfaceTemplate,
                                   ByVal SectionID As Integer,
                                   ByVal SourceWorksheets As IEnumerable(Of String))

            If Owner Is Nothing OrElse SourceWorksheets Is Nothing Then Return

            SyncLock SyncRoot

                RemoveSectionInternal(Owner, SectionID)

                Dim AddedSheets As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each WorksheetName As String In SourceWorksheets

                    If String.IsNullOrWhiteSpace(WorksheetName) Then Continue For
                    If Not AddedSheets.Add(WorksheetName) Then Continue For

                    Dim Entries As List(Of SectionDependency) = Nothing

                    If Not WorksheetSections.TryGetValue(WorksheetName, Entries) Then
                        Entries = New List(Of SectionDependency)
                        WorksheetSections.Add(WorksheetName, Entries)
                    End If

                    Entries.Add(New SectionDependency With {
                        .Owner = New WeakReference(Owner),
                        .SectionID = SectionID
                    })

                Next

                PurgeDeadReferencesInternal()

            End SyncLock

        End Sub

        Public Sub UnregisterSection(ByVal Owner As DataInterfaceTemplate,
                                     ByVal SectionID As Integer)

            If Owner Is Nothing Then Return

            SyncLock SyncRoot
                RemoveSectionInternal(Owner, SectionID)
                PurgeDeadReferencesInternal()
            End SyncLock

        End Sub

        Public Sub UnregisterInterface(ByVal Owner As DataInterfaceTemplate)

            If Owner Is Nothing Then Return

            SyncLock SyncRoot

                For Each WorksheetName As String In WorksheetSections.Keys.ToList()

                    Dim Entries As List(Of SectionDependency) = WorksheetSections(WorksheetName)

                    Entries.RemoveAll(
                        Function(Entry As SectionDependency)
                            If Entry Is Nothing OrElse Entry.Owner Is Nothing OrElse Not Entry.Owner.IsAlive Then Return True
                            Return Object.ReferenceEquals(TryCast(Entry.Owner.Target, DataInterfaceTemplate), Owner)
                        End Function)

                    If Entries.Count = 0 Then WorksheetSections.Remove(WorksheetName)

                Next

            End SyncLock

        End Sub

        Public Sub WorksheetStructureChanged(ByVal WorksheetName As String)

            If String.IsNullOrWhiteSpace(WorksheetName) Then Return

            Dim Pending As New List(Of SectionDependency)

            SyncLock SyncRoot

                Dim Entries As List(Of SectionDependency) = Nothing

                If WorksheetSections.TryGetValue(WorksheetName, Entries) Then
                    Pending.AddRange(Entries)
                End If

                PurgeDeadReferencesInternal()

            End SyncLock

            'Do not hold the registry lock while an interface disposes/rebuilds controls.
            For Each Entry As SectionDependency In Pending

                If Entry Is Nothing OrElse Entry.Owner Is Nothing OrElse Not Entry.Owner.IsAlive Then Continue For

                Dim Owner As DataInterfaceTemplate = TryCast(Entry.Owner.Target, DataInterfaceTemplate)

                If Owner Is Nothing OrElse Owner.IsDisposed Then Continue For

                Owner.InvalidateSectionFromDependency(Entry.SectionID, WorksheetName)

            Next

        End Sub

        Public Function GetNamedRangeWorksheetName(ByVal NamedRange As String) As String

            If String.IsNullOrWhiteSpace(NamedRange) Then Return Nothing
            If FileManager.ExcelModels Is Nothing Then Return Nothing
            If ModelID < 0 OrElse ModelID >= FileManager.ExcelModels.Length Then Return Nothing
            If FileManager.ExcelModels(ModelID) Is Nothing Then Return Nothing

            Dim WB As IWorkbook = FileManager.ExcelModels(ModelID).WB

            If WB Is Nothing Then Return Nothing

            'First try a workbook-level defined name.
            Try
                Dim DN As DefinedName = WB.DefinedNames.GetDefinedName(NamedRange)
                If DN IsNot Nothing AndAlso DN.Range IsNot Nothing AndAlso DN.Range.Worksheet IsNot Nothing Then
                    Return DN.Range.Worksheet.Name
                End If
            Catch
            End Try

            'Then allow for a worksheet-local defined name.
            For Each WS As Worksheet In WB.Worksheets
                Try
                    Dim DN As DefinedName = WS.DefinedNames.GetDefinedName(NamedRange)
                    If DN IsNot Nothing AndAlso DN.Range IsNot Nothing Then Return WS.Name
                Catch
                End Try
            Next

            Return Nothing

        End Function

        Public Sub Clear()
            SyncLock SyncRoot
                WorksheetSections.Clear()
            End SyncLock
        End Sub

        Private Sub RemoveSectionInternal(ByVal Owner As DataInterfaceTemplate,
                                          ByVal SectionID As Integer)

            For Each WorksheetName As String In WorksheetSections.Keys.ToList()

                Dim Entries As List(Of SectionDependency) = WorksheetSections(WorksheetName)

                Entries.RemoveAll(
                    Function(Entry As SectionDependency)
                        If Entry Is Nothing OrElse Entry.Owner Is Nothing OrElse Not Entry.Owner.IsAlive Then Return True

                        Dim ExistingOwner As DataInterfaceTemplate = TryCast(Entry.Owner.Target, DataInterfaceTemplate)

                        Return ExistingOwner Is Nothing OrElse
                               (Object.ReferenceEquals(ExistingOwner, Owner) AndAlso Entry.SectionID = SectionID)
                    End Function)

                If Entries.Count = 0 Then WorksheetSections.Remove(WorksheetName)

            Next

        End Sub

        Private Sub PurgeDeadReferencesInternal()

            For Each WorksheetName As String In WorksheetSections.Keys.ToList()

                Dim Entries As List(Of SectionDependency) = WorksheetSections(WorksheetName)

                Entries.RemoveAll(
                    Function(Entry As SectionDependency)
                        Return Entry Is Nothing OrElse Entry.Owner Is Nothing OrElse Not Entry.Owner.IsAlive OrElse Entry.Owner.Target Is Nothing
                    End Function)

                If Entries.Count = 0 Then WorksheetSections.Remove(WorksheetName)

            Next

        End Sub

    End Class

End Namespace
