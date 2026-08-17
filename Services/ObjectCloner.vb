Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Namespace Abovo
    Public Class ObjectCloner
        Function DeepClone(Of Thing)(ByRef orig As Thing) As Thing

            If (Object.ReferenceEquals(orig, Nothing)) Then Return Nothing

            Dim formatter As New BinaryFormatter()
            Dim stream As New MemoryStream()

            formatter.Serialize(stream, orig)
            stream.Seek(0, SeekOrigin.Begin)

            Return CType(formatter.Deserialize(stream), Thing)

        End Function

    End Class

End Namespace
