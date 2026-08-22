Imports System.Collections.Generic

Namespace Abovo

    Public NotInheritable Class ModelResourceRegistry

        Private ReadOnly Entries As New Dictionary(Of String, ResourceEntry)(
            StringComparer.OrdinalIgnoreCase)
        Private ReadOnly SyncRoot As New Object

        Public Sub RegisterExclusive(ByVal resourceKey As String,
                                     ByVal owner As Object,
                                     ByVal releaseOwner As Action)

            ValidateRegistration(resourceKey, owner, releaseOwner)

            SyncLock SyncRoot
                If Entries.ContainsKey(resourceKey) Then
                    Throw New InvalidOperationException(
                        "The model resource '" & resourceKey &
                        "' already has an owner.")
                End If

                Entries.Add(
                    resourceKey,
                    New ResourceEntry(owner, releaseOwner))
            End SyncLock

        End Sub

        Public Sub ReleaseCurrent(ByVal resourceKey As String)

            If String.IsNullOrWhiteSpace(resourceKey) Then Return

            Dim entry As ResourceEntry = Nothing

            SyncLock SyncRoot
                If Entries.TryGetValue(resourceKey, entry) Then
                    Entries.Remove(resourceKey)
                End If
            End SyncLock

            If entry IsNot Nothing Then entry.ReleaseOwner()

        End Sub

        Public Sub Release(ByVal resourceKey As String, ByVal owner As Object)

            If String.IsNullOrWhiteSpace(resourceKey) OrElse owner Is Nothing Then Return

            SyncLock SyncRoot
                Dim entry As ResourceEntry = Nothing
                If Entries.TryGetValue(resourceKey, entry) AndAlso
                   Object.ReferenceEquals(entry.Owner, owner) Then

                    Entries.Remove(resourceKey)
                End If
            End SyncLock

        End Sub

        Public Function IsOwnedBy(ByVal resourceKey As String,
                                  ByVal owner As Object) As Boolean

            If String.IsNullOrWhiteSpace(resourceKey) OrElse owner Is Nothing Then
                Return False
            End If

            SyncLock SyncRoot
                Dim entry As ResourceEntry = Nothing
                Return Entries.TryGetValue(resourceKey, entry) AndAlso
                       Object.ReferenceEquals(entry.Owner, owner)
            End SyncLock

        End Function

        Public Sub ReleaseAll()

            Dim entriesToRelease As List(Of ResourceEntry)

            SyncLock SyncRoot
                entriesToRelease = Entries.Values.ToList()
                Entries.Clear()
            End SyncLock

            Dim failures As New List(Of Exception)

            For Each entry As ResourceEntry In entriesToRelease
                Try
                    entry.ReleaseOwner()
                Catch ex As Exception
                    failures.Add(ex)
                End Try
            Next

            If failures.Count > 0 Then
                Throw New AggregateException(
                    "One or more model resources could not be released.",
                    failures)
            End If

        End Sub

        Private Shared Sub ValidateRegistration(ByVal resourceKey As String,
                                                ByVal owner As Object,
                                                ByVal releaseOwner As Action)

            If String.IsNullOrWhiteSpace(resourceKey) Then
                Throw New ArgumentException(
                    "A model resource key is required.",
                    NameOf(resourceKey))
            End If

            If owner Is Nothing Then
                Throw New ArgumentNullException(NameOf(owner))
            End If

            If releaseOwner Is Nothing Then
                Throw New ArgumentNullException(NameOf(releaseOwner))
            End If

        End Sub

        Private NotInheritable Class ResourceEntry

            Public ReadOnly Owner As Object
            Private ReadOnly ReleaseAction As Action

            Public Sub New(ByVal setOwner As Object,
                           ByVal setReleaseAction As Action)

                Owner = setOwner
                ReleaseAction = setReleaseAction

            End Sub

            Public Sub ReleaseOwner()
                ReleaseAction.Invoke()
            End Sub

        End Class

    End Class

    Public NotInheritable Class ModelResourceKeys

        Private Sub New()
        End Sub

        Public Const TransactionalRecordsRangeDataSource As String =
            "RangeDataSource:Transactional DB!Transactional_Records"

    End Class

End Namespace
