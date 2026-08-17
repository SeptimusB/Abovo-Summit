Imports Abovo.LogDebugDev
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Imports DevExpress.Spreadsheet.Functions
Imports System.Globalization
Imports Abovo.FileManager
Imports DevExpress.XtraRichEdit.Model

Namespace Abovo
#Region "#ResposiveCostImplementation"
    Public Class ResponsiveCostFunction 'TestArrayCustomFunction
        Implements ICustomFunction

        Private Const functionName As String = "RESPCOST"
        Private ReadOnly functionParameters() As ParameterInfo

        Public Sub New()
            '(Year As Integer, AllUnits As Integer, FirstManage As Integer, LastManage As Integer, FinalYear, ApplRates As Range, ApplYears As Range)
            ' Missing optional parameters do not result in an error message.

            Me.functionParameters = New ParameterInfo() {
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Optional),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Optional),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Optional),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Optional),
            New ParameterInfo(ParameterType.Reference, ParameterAttributes.Optional),
            New ParameterInfo(ParameterType.Reference, ParameterAttributes.Optional)
        }

        End Sub
        Public ReadOnly Property Name() As String Implements IFunction.Name

            Get

                Return functionName

            End Get

        End Property
        Private ReadOnly Property IFunction_Parameters() As ParameterInfo() Implements IFunction.Parameters

            Get

                Return functionParameters

            End Get

        End Property
        Private ReadOnly Property IFunction_ReturnType() As ParameterType Implements IFunction.ReturnType

            Get

                Return ParameterType.Value

            End Get

        End Property
        Public ReadOnly Property Volatile() As Boolean Implements IFunction.Volatile

            Get

                Return False

            End Get

        End Property
        Public Function GetName(ByVal culture As CultureInfo) As String Implements ICustomFunction.GetName

            Return Name

        End Function
        Public Function IFunction_Evaluate(ByVal parameters As IList(Of ParameterValue), ByVal context As EvaluationContext) As ParameterValue Implements IFunction.Evaluate

            ''''''Original VBA Function
            'Function PMCost(Year As Integer, AllUnits As Integer, UnitCost As Double, FirstManage As Integer, LastManage As Integer, FinalYear, ApplRates As Range, ApplYears As Range)
            'Function RespCost(Year As Integer, AllUnits As Integer, FirstManage As Integer, LastManage As Integer, FinalYear, ApplRates As Range, ApplYears As Range)
            '            Application.Volatile(False)

            '            intResCount = intResCount + 1

            '            'function to calculate planned maintenance costs



            '            Dim AnnualUnits As Double
            '            Dim AnnualRate As Double
            '            Dim TotCost As Double
            '            Dim i As Integer, j As Integer

            '            On Error Resume Next

            '            TotCost = 0

            '            AnnualUnits = AllUnits / (LastManage - FirstManage + 1)

            '            AnnualRate = 0

            '            i = Year

            '            j = 0

            '            Do While i >= FirstManage And j < (LastManage - FirstManage + 1)

            '                AnnualRate = ApplRates.Cells(Application.WorksheetFunction.Match(i - FirstManage + 1, ApplYears))
            '                TotCost = TotCost + AnnualUnits * AnnualRate
            '                i = i - 1
            '                j = j + 1

            '            Loop

            '            RespCost = TotCost



            '        End Function




            Dim Engine As FormulaEngine = context.Sheet.Workbook.FormulaEngine '
            Dim IntYear As Integer = Convert.ToInt32(parameters(0).NumericValue)
            Dim AllUnits As Integer = Convert.ToInt32(parameters(1).NumericValue)
            Dim FirstManage As Integer = Convert.ToInt32(parameters(2).NumericValue)
            Dim LastManage As Integer = Convert.ToInt32(parameters(3).NumericValue)
            Dim FinalYear As Integer = Convert.ToInt32(parameters(4).NumericValue) ' This is unused but remains for compatibility
            Dim ApplRates As CellRange = parameters(5).RangeValue
            Dim ApplYears As CellRange = parameters(6).RangeValue

            Dim AnnualUnits As Double = AllUnits / (LastManage - FirstManage + 1)
            Dim AnnualRate As Double = 0
            Dim TotCost As Double = 0
            Dim i As Integer = IntYear
            Dim j As Integer = 0
            Dim ValReturn As ParameterValue
            Dim StrMatch As String = ""
            Dim expcontext As New ExpressionContext(context.Column, context.Row, context.Sheet, context.Culture, ReferenceStyle.R1C1, DevExpress.Spreadsheet.Formulas.ExpressionStyle.Normal)
            Do While i >= FirstManage And j < (LastManage - FirstManage + 1)

                AnnualRate = 0
                StrMatch = "=MATCH(" & i - FirstManage + 1 & ", " & ApplYears.GetReferenceR1C1(ReferenceElement.IncludeSheetName Or ReferenceElement.RowAbsolute Or ReferenceElement.ColumnAbsolute, Nothing) & ")"
                SystemLog(StrMatch)

                ValReturn = Engine.Evaluate(StrMatch, expcontext)
                'If IsNumeric(ValReturn) Then
                AnnualRate = ApplRates(ValReturn.NumericValue - 1).Value.NumericValue
                'Else
                '    Debug.Print("RC Error " & context.Sheet.ToString & " R" & context.Row.ToString & "C" & context.Column.ToString & StrMatch & " " & ValReturn.ToString)
                '    AnnualRate = 0 ' If no match, set rate to 0
                'End If
                TotCost += AnnualUnits * AnnualRate
                i -= 1
                j += 1

            Loop

            Return TotCost
            systemlog("RespCost return " & TotCost.ToString)

        End Function

    End Class
#End Region ' #ResposiveCostImplementation
End Namespace