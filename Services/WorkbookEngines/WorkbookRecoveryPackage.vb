Option Strict On

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Namespace Abovo.WorkbookEngines
    Partial Friend NotInheritable Class WorkbookCandidatePackage
        Friend Sub RequireRecoveryProfile(source As String)
            If parts.ContainsKey("xl/metadata.bin") AndAlso Not RecoveryXlsmCompatibility.HasVerifiedBinaryProfile(source) Then
                Throw New NotSupportedException("This workbook's dynamic-array metadata needs a separately verified XLSM conversion. The previous recovery copy is unchanged.")
            End If
        End Sub

        Friend Sub VerifyRecoveryCandidate(path As String, source As String)
            RequireRecoveryProfile(source)
            Dim expected As New Dictionary(Of String, String)(parts, StringComparer.Ordinal)
            If expected.ContainsKey("xl/metadata.bin") Then
                expected.Remove("xl/metadata.bin")
                Dim canonicalText As New StringBuilder()
                Canonical(RecoveryXlsmCompatibility.VerifiedMetadataDocument().Root, canonicalText)
                Using hash = SHA256.Create()
                    expected.Add("xl/metadata.xml", Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(canonicalText.ToString()))))
                End Using
            End If
            Dim actual = Capture(path)
            Dim normalise As Func(Of String, String) = Function(link) link.Replace("xl/workbook.bin|", "xl/workbook.xml|").Replace("|xl/metadata.bin", "|xl/metadata.xml")
            If expected.Count <> actual.parts.Count OrElse
                Not New HashSet(Of String)(links.Select(normalise), StringComparer.Ordinal).SetEquals(actual.links) Then
                Throw New InvalidDataException("Recovery conversion changed the custom XML, VBA or dynamic-array package structure.")
            End If
            For Each pair In expected
                Dim actualHash As String = Nothing
                If Not actual.parts.TryGetValue(pair.Key, actualHash) OrElse actualHash <> pair.Value Then
                    Throw New InvalidDataException("Recovery conversion did not preserve " & pair.Key & ". The previous recovery copy is unchanged.")
                End If
            Next
        End Sub
    End Class
End Namespace
