Option Strict On

Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.Win32.SafeHandles

Namespace Abovo.WorkbookEngines
    ' Local NTFS qualification only. A path hash followed by File.Replace leaves
    ' a pathname race. Rename the checked handle, never an unchecked path, and
    ' never replace a destination which appeared during the hand-off.
    Friend NotInheritable Class WorkbookPublicationFile
        Implements IDisposable
        Private ReadOnly handle As SafeFileHandle
        Private ReadOnly stream As FileStream

        <StructLayout(LayoutKind.Sequential)>
        Private Structure NativeFileTime
            Public Low As UInteger
            Public High As UInteger
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure FileInformation
            Public Attributes As UInteger
            Public Creation As NativeFileTime
            Public Access As NativeFileTime
            Public Write As NativeFileTime
            Public Volume As UInteger
            Public SizeHigh As UInteger
            Public SizeLow As UInteger
            Public Links As UInteger
            Public IndexHigh As UInteger
            Public IndexLow As UInteger
        End Structure

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure StreamInformation
            Public Size As Long
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=296)>
            Public Name As String
        End Structure

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, EntryPoint:="CreateFileW")>
        Private Shared Function OpenNative(path As String, access As UInteger, share As UInteger, security As IntPtr,
                                           creation As UInteger, flags As UInteger, template As IntPtr) As SafeFileHandle
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function GetFileInformationByHandle(file As SafeFileHandle, ByRef info As FileInformation) As Boolean
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function SetFileInformationByHandle(file As SafeFileHandle, kind As Integer, info As IntPtr, size As UInteger) As Boolean
        End Function
        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function FindFirstStreamW(path As String, level As Integer, ByRef info As StreamInformation, flags As UInteger) As IntPtr
        End Function
        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function FindNextStreamW(search As IntPtr, ByRef info As StreamInformation) As Boolean
        End Function
        <DllImport("kernel32.dll")>
        Private Shared Function FindClose(search As IntPtr) As Boolean
        End Function

        Friend Sub New(path As String, allowRename As Boolean)
            handle = OpenNative(path, &H80000000UI Or If(allowRename, &H10000UI, 0UI), 1UI,
                                IntPtr.Zero, 3UI, &H200000UI, IntPtr.Zero)
            If handle.IsInvalid Then
                Dim code = Marshal.GetLastWin32Error()
                handle.Dispose()
                Throw New IOException("Cannot exclusively guard workbook file: " & path, New Win32Exception(code))
            End If
            Try
                Dim info As FileInformation
                If Not GetFileInformationByHandle(handle, info) Then Throw New Win32Exception(Marshal.GetLastWin32Error())
                If info.Links <> 1 OrElse (info.Attributes And &H5414UI) <> 0 Then Throw New IOException("Linked, redirected, encrypted, offline or system files are outside this publication trial.")
                If allowRename AndAlso (info.Attributes And 1UI) <> 0 Then Throw New IOException("A read-only file cannot be replaced.")
                RequireSupportedStreams(path)
                stream = New FileStream(handle, FileAccess.Read)
            Catch
                handle.Dispose()
                Throw
            End Try
        End Sub

        Friend Function Hash() As String
            stream.Position = 0
            Using algorithm = SHA256.Create()
                Return BitConverter.ToString(algorithm.ComputeHash(stream)).Replace("-", "")
            End Using
        End Function

        Friend Sub RenameNew(path As String)
            Dim bytes = Encoding.Unicode.GetBytes(path)
            Dim rootOffset = If(IntPtr.Size = 8, 8, 4)
            Dim lengthOffset = rootOffset + IntPtr.Size
            Dim nameOffset = lengthOffset + 4
            Dim size = nameOffset + bytes.Length + 2
            Dim info = Marshal.AllocHGlobal(size)
            Try
                Marshal.Copy(New Byte(size - 1) {}, 0, info, size)
                ' ReplaceIfExists=False, RootDirectory=NULL, fully qualified name.
                Marshal.WriteInt32(info, lengthOffset, bytes.Length)
                Marshal.Copy(bytes, 0, IntPtr.Add(info, nameOffset), bytes.Length)
                If Not SetFileInformationByHandle(handle, 3, info, CUInt(size)) Then Throw New IOException("Workbook rename was refused; no destination was overwritten.", New Win32Exception(Marshal.GetLastWin32Error()))
            Finally
                Marshal.FreeHGlobal(info)
            End Try
        End Sub

        Friend Shared Function Normalize(path As String) As String
            If String.IsNullOrWhiteSpace(path) OrElse path.Length < 3 OrElse path(1) <> ":"c OrElse
                (path(2) <> "\"c AndAlso path(2) <> "/"c) OrElse path.Substring(2).Contains(":") Then Throw New ArgumentException("A local absolute workbook path is required.")
            Dim full = IO.Path.GetFullPath(path)
            If full.Length >= 248 Then Throw New ArgumentException("Long paths are outside this publication trial.")
            For Each part In full.Substring(3).Split("\"c)
                If part.Length = 0 OrElse part.EndsWith(".", StringComparison.Ordinal) OrElse part.EndsWith(" ", StringComparison.Ordinal) OrElse
                    Text.RegularExpressions.Regex.IsMatch(part, "^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])(\.|$)", Text.RegularExpressions.RegexOptions.IgnoreCase) Then Throw New ArgumentException("Unsupported workbook path component.")
            Next
            Dim drive As New DriveInfo(IO.Path.GetPathRoot(full))
            If drive.DriveType <> DriveType.Fixed OrElse Not String.Equals(drive.DriveFormat, "NTFS", StringComparison.OrdinalIgnoreCase) Then Throw New NotSupportedException("Publication trial requires a local fixed NTFS volume; network/cloud paths need separate qualification.")
            If Not Directory.Exists(IO.Path.GetDirectoryName(full)) Then Throw New DirectoryNotFoundException("Destination folder does not exist.")
            Return full
        End Function

        Friend Shared Function PinDirectories(paths As IEnumerable(Of String)) As List(Of SafeFileHandle)
            Dim pins As New List(Of SafeFileHandle)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Try
                For Each path In paths
                    Dim chain As New Stack(Of DirectoryInfo)()
                    Dim directory As New DirectoryInfo(IO.Path.GetDirectoryName(Normalize(path)))
                    While directory IsNot Nothing
                        chain.Push(directory) : directory = directory.Parent
                    End While
                    For Each item In chain
                        If Not seen.Add(item.FullName) Then Continue For
                        Dim pin = OpenNative(item.FullName, 0UI, 3UI, IntPtr.Zero, 3UI, &H2200000UI, IntPtr.Zero)
                        If pin.IsInvalid Then
                            Dim code = Marshal.GetLastWin32Error() : pin.Dispose()
                            Throw New IOException("Cannot guard publication directory.", New Win32Exception(code))
                        End If
                        pins.Add(pin)
                        Dim info As FileInformation
                        If Not GetFileInformationByHandle(pin, info) Then Throw New Win32Exception(Marshal.GetLastWin32Error())
                        If (info.Attributes And &H5400UI) <> 0 Then Throw New IOException("Redirected, offline or encrypted directories are outside this publication trial.")
                    Next
                Next
                Return pins
            Catch
                For Each pin In pins
                    pin.Dispose()
                Next
                Throw
            End Try
        End Function

        Friend Shared Sub RequireSupportedStreams(path As String)
            StreamNames(path)
        End Sub

        Private Shared Function StreamNames(path As String) As List(Of String)
            Dim names As New List(Of String)()
            Dim info As New StreamInformation()
            Dim search = FindFirstStreamW(path, 0, info, 0)
            If search = New IntPtr(-1) Then Throw New IOException("Unable to inspect workbook streams.", New Win32Exception(Marshal.GetLastWin32Error()))
            Try
                Do
                    If Not String.Equals(info.Name, "::$DATA", StringComparison.OrdinalIgnoreCase) AndAlso
                        Not String.Equals(info.Name, ":Zone.Identifier:$DATA", StringComparison.OrdinalIgnoreCase) AndAlso
                        Not String.Equals(info.Name, ":MBAM.Zone.Identifier:$DATA", StringComparison.OrdinalIgnoreCase) Then Throw New IOException("Workbook has an unsupported alternate stream: " & info.Name)
                    If info.Name <> "::$DATA" Then names.Add(info.Name)
                    If Not FindNextStreamW(search, info) Then
                        Dim code = Marshal.GetLastWin32Error()
                        If code <> 38 Then Throw New IOException("Workbook stream inspection failed.", New Win32Exception(code))
                        Exit Do
                    End If
                Loop
            Finally
                FindClose(search)
            End Try
            Return names
        End Function

        ' Preserve recognized provenance markers as opaque bytes. In particular,
        ' MBAM.Zone.Identifier is NOT interpreted as macro permission or trust.
        Friend Shared Function CaptureMarkers(path As String) As Dictionary(Of String, Byte())
            Dim markers As New Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)
            For Each name In StreamNames(path)
                Using marker = OpenNative(path & name, &H80000000UI, 7UI, IntPtr.Zero, 3UI, 0UI, IntPtr.Zero)
                    If marker.IsInvalid Then Throw New IOException("Cannot read workbook provenance marker.", New Win32Exception(Marshal.GetLastWin32Error()))
                    Using input As New FileStream(marker, FileAccess.Read)
                        If input.Length > 65536 Then Throw New IOException("Workbook provenance marker exceeds trial limits.")
                        Using reader As New BinaryReader(input)
                            Dim data = reader.ReadBytes(65537)
                            If data.Length > 65536 Then Throw New IOException("Workbook provenance marker changed beyond trial limits.")
                            markers.Add(name, data)
                        End Using
                    End Using
                End Using
            Next
            Return markers
        End Function

        Friend Shared Function MarkersMatch(path As String, expected As Dictionary(Of String, Byte())) As Boolean
            Dim actual = CaptureMarkers(path)
            Return actual.Count = expected.Count AndAlso expected.All(Function(pair) actual.ContainsKey(pair.Key) AndAlso actual(pair.Key).SequenceEqual(pair.Value))
        End Function

        Friend Shared Sub WriteMarkers(path As String, markers As Dictionary(Of String, Byte()))
            For Each pair In markers
                Using marker = OpenNative(path & pair.Key, &H40000000UI, 1UI, IntPtr.Zero, 1UI, 0UI, IntPtr.Zero)
                    If marker.IsInvalid Then Throw New IOException("Cannot preserve workbook provenance marker.", New Win32Exception(Marshal.GetLastWin32Error()))
                    Using output As New FileStream(marker, FileAccess.Write)
                        output.Write(pair.Value, 0, pair.Value.Length) : output.Flush(True)
                    End Using
                End Using
            Next
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If stream IsNot Nothing Then stream.Dispose()
            handle.Dispose()
        End Sub
    End Class
End Namespace
