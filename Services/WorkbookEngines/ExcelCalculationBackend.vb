Option Strict Off

Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Microsoft.Win32.SafeHandles

Namespace Abovo.WorkbookEngines
    ' Late-bound out-of-process COM supports either installed Office bitness
    ' without adding an Office PIA/package or taking ownership of user windows.
    Partial Friend NotInheritable Class ExcelCalculationBackend
        Implements IWorkbookCalculationBackend

        Private ReadOnly threadId As Integer = Thread.CurrentThread.ManagedThreadId
        Private app As Object
        Private books As Object
        Private seed As Object
        Private book As Object
        Private owned As Boolean
        Private versionText As String
        Private filter As ExcelBusyCallFilter

        <DllImport("user32.dll")>
        Private Shared Function GetWindowThreadProcessId(window As IntPtr, ByRef processId As UInteger) As UInteger
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, EntryPoint:="CreateFileW")>
        Private Shared Function OpenNativeFile(path As String, access As UInteger, share As UInteger,
                                               security As IntPtr, creation As UInteger, flags As UInteger,
                                               template As IntPtr) As SafeFileHandle
        End Function

        Public ReadOnly Property Name As String Implements IWorkbookCalculationBackend.Name
            Get
                Return "Excel"
            End Get
        End Property
        Public ReadOnly Property Version As String Implements IWorkbookCalculationBackend.Version
            Get
                Return versionText
            End Get
        End Property
        Private Sub RequireOwner()
            If Thread.CurrentThread.ManagedThreadId <> threadId OrElse Thread.CurrentThread.GetApartmentState() <> ApartmentState.STA Then
                Throw New InvalidOperationException("Excel accessed outside its STA owner.")
            End If
        End Sub

        Private Shared Sub InspectPackage(path As String)
            ' Do not trust a temporary path to bypass downloaded-file policy.
            Dim zone = ReadInternetZone(path)
            If zone IsNot Nothing Then
                If Text.RegularExpressions.Regex.IsMatch(zone, "(?im)^\s*ZoneId\s*=\s*[34]\s*$") Then
                    Throw New InvalidOperationException("Excel calculation is unavailable for a workbook marked as downloaded or restricted. Its security marking has not been changed.")
                End If
            End If
            Using archive = ZipFile.OpenRead(path)
                For Each entry In archive.Entries
                    Dim name = entry.FullName
                    If name.StartsWith("xl/macrosheets/", StringComparison.OrdinalIgnoreCase) OrElse
                       name.StartsWith("xl/intlmacrosheets/", StringComparison.OrdinalIgnoreCase) Then
                        Throw New InvalidOperationException("Excel 4 macro sheets are not supported by this calculation adapter.")
                    End If
                    If name.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(name, "xl/connections.xml", StringComparison.OrdinalIgnoreCase) Then
                        Throw New InvalidOperationException("External workbook links/connections require a separately reviewed Excel opening policy.")
                    End If
                Next
            End Using
        End Sub

        Private Shared Function ReadInternetZone(path As String) As String
            ' .NET Framework rejects the colon in an alternate-stream path.
            ' Use a read-only native handle; never remove or replace the marker.
            Using handle = OpenNativeFile(path & ":Zone.Identifier", &H80000000UI, 7UI, IntPtr.Zero, 3UI, 0UI, IntPtr.Zero)
                If handle.IsInvalid Then
                    Dim errorCode = Marshal.GetLastWin32Error()
                    If errorCode = 2 Then Return Nothing ' Stream does not exist.
                    Throw New ComponentModel.Win32Exception(errorCode, "The workbook security marker could not be inspected.")
                End If
                Using stream As New FileStream(handle, FileAccess.Read), reader As New StreamReader(stream)
                    Return reader.ReadToEnd()
                End Using
            End Using
        End Function

        Public Sub OpenReadOnly(path As String, options As WorkbookEngineOptions) Implements IWorkbookCalculationBackend.OpenReadOnly
            RequireOwner()
            If options.RequireSummitFunctions AndAlso Not options.AllowTrustedVba Then Throw New InvalidOperationException("Required workbook VBA functions are disabled by the selected policy.")
            InspectPackage(path)
            Dim excelType = Type.GetTypeFromProgID("Excel.Application", False)
            If excelType Is Nothing Then Throw New InvalidOperationException("Compatible desktop Excel is not installed.")
            Dim existing As New HashSet(Of Integer)()
            For Each existingProcess In Process.GetProcessesByName("EXCEL")
                Using existingProcess
                    existing.Add(existingProcess.Id)
                End Using
            Next
            filter = New ExcelBusyCallFilter()
            app = Activator.CreateInstance(excelType)
            Dim processId As UInteger
            GetWindowThreadProcessId(New IntPtr(CInt(app.Hwnd)), processId)
            If processId = 0 OrElse existing.Contains(CInt(processId)) Then Throw New InvalidOperationException("Refusing to use a pre-existing Excel process.")
            owned = True
            app.AutomationSecurity = 3 ' ForceDisable until the explicit ByUI open below.
            app.EnableEvents = False : app.Visible = False : app.DisplayAlerts = False
            app.AskToUpdateLinks = False : app.ScreenUpdating = False
            versionText = "Excel " & CStr(app.Version) & " build " & CStr(app.Build)
            If Double.Parse(CStr(app.Version), CultureInfo.InvariantCulture) < 16 Then Throw New InvalidOperationException("This Excel version does not meet the initial compatibility policy.")
            books = app.Workbooks
            If CInt(books.Count) <> 0 Then Throw New InvalidOperationException("Excel opened an unexpected startup workbook.")
            seed = books.Add()
            app.Calculation = -4135 : app.CalculateBeforeSave = False
            VerifyDynamicArrays()
            Try
                app.AutomationSecurity = If(options.AllowTrustedVba, 2, 3) ' ByUI; never Low/force-enable.
                book = books.Open(path, 0, True) ' Do not update links; original is read-only.
            Finally
                app.AutomationSecurity = 3
            End Try
            app.Calculation = -4135
            If Not CBool(book.ReadOnly) Then Throw New InvalidOperationException("Excel did not open the workbook read-only.")
            If options.RequireSummitFunctions Then VerifySummitFunctions()
        End Sub

        Private Sub VerifyDynamicArrays()
            Dim sheets As Object = Nothing, sheet As Object = Nothing, anchor As Object = Nothing, spill As Object = Nothing
            Try
                sheets = seed.Worksheets : sheet = sheets.Item(1)
                anchor = sheet.Range("D1") : spill = sheet.Range("D2")
                anchor.Formula2 = "=SEQUENCE(2)"
                sheet.Calculate()
                If Not TypeOf spill.Value2 Is Double OrElse CDbl(spill.Value2) <> 2 Then Throw New InvalidOperationException("Excel dynamic-array capability check failed.")
                anchor.ClearContents()
            Finally
                Release(spill) : Release(anchor) : Release(sheet) : Release(sheets)
            End Try
        End Sub

        Private Sub VerifySummitFunctions()
            Dim sheets As Object = Nothing, sheet As Object = Nothing, rates As Object = Nothing, years As Object = Nothing
            Try
                sheets = seed.Worksheets : sheet = sheets.Item(1)
                rates = sheet.Range("A1:A40") : years = sheet.Range("B1:B40")
                Dim rv(39, 0) As Object, yv(39, 0) As Object
                For i As Integer = 0 To 39
                    rv(i, 0) = (i + 1) * 0.02 : yv(i, 0) = CDbl(i + 1)
                Next
                rates.Value2 = rv : years.Value2 = yv
                Dim prefix = "'" & CStr(book.Name).Replace("'", "''") & "'!"
                Dim planned = app.Run(prefix & "PMCost", 10, 100, 5.0, 1, 5, 40, rates, years)
                Dim responsive = app.Run(prefix & "RespCost", 10, 100, 1, 5, 40, rates, years)
                If Not TypeOf planned Is Double OrElse Not TypeOf responsive Is Double OrElse
                   Math.Abs(CDbl(planned) - 80) > 0.000001 OrElse Math.Abs(CDbl(responsive) - 16) > 0.000001 Then
                    Throw New InvalidOperationException("Workbook VBA function compatibility check failed.")
                End If
            Finally
                Release(years) : Release(rates) : Release(sheet) : Release(sheets)
            End Try
        End Sub

        Public Sub Calculate(kind As WorkbookCalculationKind) Implements IWorkbookCalculationBackend.Calculate
            RequireOwner()
            Select Case kind
                Case WorkbookCalculationKind.Incremental : app.Calculate()
                Case WorkbookCalculationKind.Full : app.CalculateFull()
                Case WorkbookCalculationKind.Rebuild : app.CalculateFullRebuild()
                Case Else : Throw New ArgumentOutOfRangeException(NameOf(kind))
            End Select
            If CInt(app.CalculationState) <> 0 Then Throw New InvalidOperationException("Excel has not finished calculating; no results were published.")
        End Sub

        Public Function Read(area As WorkbookReadArea) As WorkbookValueBlock Implements IWorkbookCalculationBackend.Read
            RequireOwner()
            Dim sheets As Object = Nothing, sheet As Object = Nothing, range As Object = Nothing
            Try
                sheets = book.Worksheets : sheet = sheets.Item(area.Worksheet) : range = sheet.Range(area.Address)
                Dim raw As Object = range.Value2
                Dim matrix = TryCast(raw, Object(,))
                Dim data(area.Rows - 1, area.Columns - 1) As Object
                For r As Integer = 0 To area.Rows - 1
                    For c As Integer = 0 To area.Columns - 1
                        data(r, c) = ConvertValue(If(matrix Is Nothing, raw, matrix(r + matrix.GetLowerBound(0), c + matrix.GetLowerBound(1))))
                    Next
                Next
                Return New WorkbookValueBlock(area, data)
            Finally
                Release(range) : Release(sheet) : Release(sheets)
            End Try
        End Function

        Private Shared Function ConvertValue(value As Object) As Object
            If TypeOf value Is ErrorWrapper Then Return ErrorValue(DirectCast(value, ErrorWrapper).ErrorCode)
            If TypeOf value Is Integer Then
                Dim code = CInt(value)
                ' Modern errors (including #SPILL!) lie beyond the seven
                ' original CVErr values. Never turn an unknown HRESULT into a number.
                If (code And &HFFFF0000) = &H800A0000 Then Return ErrorValue(code)
                Return CDbl(code)
            End If
            Return value
        End Function

        Private Shared Function ErrorValue(code As Integer) As WorkbookCellError
            Select Case code And &HFFFF
                Case 2000 : Return New WorkbookCellError("#NULL!")
                Case 2007 : Return New WorkbookCellError("#DIV/0!")
                Case 2015 : Return New WorkbookCellError("#VALUE!")
                Case 2023 : Return New WorkbookCellError("#REF!")
                Case 2029 : Return New WorkbookCellError("#NAME?")
                Case 2036 : Return New WorkbookCellError("#NUM!")
                Case 2042 : Return New WorkbookCellError("#N/A")
                Case 2045 : Return New WorkbookCellError("#SPILL!")
                Case Else : Return New WorkbookCellError("#ERROR:" & code.ToString(CultureInfo.InvariantCulture))
            End Select
        End Function

        Private Shared Sub Release(ByRef value As Object)
            If value IsNot Nothing AndAlso Marshal.IsComObject(value) Then Marshal.ReleaseComObject(value)
            value = Nothing
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            RequireOwner()
            Dim errors As New List(Of Exception)()
            For Each item In {book, seed}
                If item Is Nothing Then Continue For
                Try
                    item.Close(False)
                Catch ex As Exception
                    errors.Add(ex)
                End Try
            Next
            Try
                If app IsNot Nothing AndAlso owned Then app.Quit()
            Catch ex As Exception
                errors.Add(ex)
            Finally
                Release(book) : Release(seed) : Release(books) : Release(app)
                If filter IsNot Nothing Then filter.Dispose()
                filter = Nothing : owned = False
            End Try
            If errors.Count > 0 Then Throw New AggregateException("Excel cleanup did not complete cleanly.", errors)
        End Sub
    End Class

    <ComImport(), Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Friend Interface IExcelOleMessageFilter
        <PreserveSig()> Function HandleInComingCall(callType As Integer, caller As IntPtr, tick As Integer, info As IntPtr) As Integer
        <PreserveSig()> Function RetryRejectedCall(callee As IntPtr, tick As Integer, rejectType As Integer) As Integer
        <PreserveSig()> Function MessagePending(callee As IntPtr, tick As Integer, pendingType As Integer) As Integer
    End Interface

    <ComVisible(True)>
    Friend NotInheritable Class ExcelBusyCallFilter
        Implements IExcelOleMessageFilter, IDisposable
        Private previous As IExcelOleMessageFilter
        <DllImport("ole32.dll")>
        Private Shared Function CoRegisterMessageFilter(filter As IExcelOleMessageFilter, ByRef prior As IExcelOleMessageFilter) As Integer
        End Function
        Public Sub New()
            Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(Me, previous))
        End Sub
        Public Function HandleInComingCall(callType As Integer, caller As IntPtr, tick As Integer, info As IntPtr) As Integer Implements IExcelOleMessageFilter.HandleInComingCall
            Return 0
        End Function
        Public Function RetryRejectedCall(callee As IntPtr, tick As Integer, rejectType As Integer) As Integer Implements IExcelOleMessageFilter.RetryRejectedCall
            Return If(rejectType = 2 AndAlso tick < 5000, 100, -1)
        End Function
        Public Function MessagePending(callee As IntPtr, tick As Integer, pendingType As Integer) As Integer Implements IExcelOleMessageFilter.MessagePending
            Return 2
        End Function
        Public Sub Dispose() Implements IDisposable.Dispose
            Dim ignored As IExcelOleMessageFilter = Nothing
            Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(previous, ignored))
            previous = Nothing
        End Sub
    End Class
End Namespace
