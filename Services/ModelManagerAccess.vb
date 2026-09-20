Imports System.Configuration
Imports System.Security.Cryptography

Namespace Abovo
    Public NotInheritable Class ModelManagerAccess
        Public Shared ReadOnly Property PasswordRequired As Boolean
            Get
                'Only an explicit false disables the gate; absent/bad configuration fails closed.
                Return Not String.Equals(ConfigurationManager.AppSettings("ModelManager.RequirePassword"), "false", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        Public Shared Function Verify(password As String, verifier As String) As Boolean
            Try
                Dim parts = verifier.Split("$"c)
                Dim iterations As Integer
                If parts.Length <> 4 OrElse parts(0) <> "pbkdf2-sha256" OrElse Not Integer.TryParse(parts(1), iterations) OrElse iterations < 200000 OrElse iterations > 2000000 Then Return False
                Dim salt = Convert.FromBase64String(parts(2))
                Dim expected = Convert.FromBase64String(parts(3))
                If salt.Length <> 16 OrElse expected.Length <> 32 Then Return False
                Using derive As New Rfc2898DeriveBytes(password, salt, iterations, HashAlgorithmName.SHA256)
                    Dim actual = derive.GetBytes(32)
                    Dim difference As Integer = 0
                    For index = 0 To 31
                        difference = difference Or (actual(index) Xor expected(index))
                    Next
                    Return difference = 0
                End Using
            Catch ex As Exception When TypeOf ex Is FormatException OrElse TypeOf ex Is ArgumentException OrElse TypeOf ex Is NullReferenceException
                Return False
            End Try
        End Function

        Public Shared Function CreateVerifier(password As String) As String
            If String.IsNullOrWhiteSpace(password) OrElse password.Length < 12 Then Throw New ArgumentException("Use at least 12 characters for the Model Manager password.")
            Dim salt(15) As Byte
            Using random = RandomNumberGenerator.Create()
                random.GetBytes(salt)
            End Using
            Using derive As New Rfc2898DeriveBytes(password, salt, 200000, HashAlgorithmName.SHA256)
                Return "pbkdf2-sha256$200000$" & Convert.ToBase64String(salt) & "$" & Convert.ToBase64String(derive.GetBytes(32))
            End Using
        End Function
    End Class
End Namespace
