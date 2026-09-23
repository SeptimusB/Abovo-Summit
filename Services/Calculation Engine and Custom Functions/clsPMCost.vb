Imports Abovo.LogDebugDev
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Formulas
Imports DevExpress.Spreadsheet.Functions
Imports System.Globalization
Imports Abovo.FileManager

Namespace Abovo
#Region "#ProjectManagememntCostImplementation"
    Public Class PMCostFunction 'TestArrayCustomFunction
        Implements ICustomFunction

        Private Const functionName As String = "PMCost"
        Private ReadOnly functionParameters() As ParameterInfo

        Public Sub New()
            ' Match the VBA signature: Year, AllUnits, UnitCost, FirstManage,
            ' LastManage, FinalYear, ApplRates, ApplYears. FinalYear is unused,
            ' but its required sixth position preserves Excel/VBA compatibility.

            Me.functionParameters = New ParameterInfo() {
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
            New ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
            New ParameterInfo(ParameterType.Reference, ParameterAttributes.Required),
            New ParameterInfo(ParameterType.Reference, ParameterAttributes.Required)
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
            'function to calculate planned maintenance costs

            'Dim AnnualUnits As Double
            'Dim AnnualRate As Double
            'Dim TotCost As Double
            'Dim i As Integer, j As Integer

            'On Error Resume Next

            'TotCost = 0
            'AnnualUnits = AllUnits / (LastManage - FirstManage + 1)
            'AnnualRate = 0
            'i = Year()
            'j = 0

            'Do While i >= FirstManage And j < (LastManage - FirstManage + 1)

            '    AnnualRate = ApplRates.Cells(Application.WorksheetFunction.Match(i - FirstManage + 1, ApplYears))
            '    TotCost = TotCost + UnitCost * AnnualUnits * AnnualRate
            '    i = i - 1
            '    j = j + 1

            'Loop

            'PMCost = TotCost

            If parameters Is Nothing OrElse parameters.Count <> 8 Then Return ParameterValue.ErrorInvalidValueInFunction
            For Each parameter In parameters
                If parameter.IsError Then Return parameter
            Next
            If Not parameters(6).IsRange OrElse Not parameters(7).IsRange Then Return ParameterValue.ErrorInvalidValueInFunction
            Dim IntYear As Integer = Convert.ToInt32(parameters(0).NumericValue)
            Dim AllUnits As Integer = Convert.ToInt32(parameters(1).NumericValue)
            Dim UnitCost As Double = Convert.ToDouble(parameters(2).NumericValue)
            Dim FirstManage As Integer = Convert.ToInt32(parameters(3).NumericValue)
            Dim LastManage As Integer = Convert.ToInt32(parameters(4).NumericValue)
            Dim FinalYear As Integer = Convert.ToInt32(parameters(5).NumericValue) ' This is unused but remains for compatibility
            Dim ApplRates As CellRange = parameters(6).RangeValue

            Dim AnnualUnits As Double = AllUnits / (LastManage - FirstManage + 1)
            Dim AnnualRate As Double = 0
            Dim TotCost As Double = 0
            Dim i As Integer = IntYear
            Dim j As Integer = 0
            Dim ValReturn As ParameterValue
            'Use DevExpress's native approximate MATCH without parsing a formula
            'and rebuilding its range address for every year. Keep references live.
            Dim match = context.Sheet.Workbook.Functions("MATCH")
            Dim matchArguments As ParameterValue() = {0, parameters(7), 1}

            Do While i >= FirstManage And j < (LastManage - FirstManage + 1)

                AnnualRate = 0
                matchArguments(0) = i - FirstManage + 1
                ValReturn = match.Evaluate(matchArguments, context)
                If ValReturn.IsError Then Return ValReturn
                If Not ValReturn.IsNumeric Then Return ParameterValue.ErrorInvalidValueInFunction
                If ValReturn.NumericValue < 1 OrElse ValReturn.NumericValue > CLng(ApplRates.RowCount) * ApplRates.ColumnCount Then Return ParameterValue.ErrorReference
                If ApplRates(ValReturn.NumericValue - 1).Value.IsError Then Return ApplRates(ValReturn.NumericValue - 1).Value
                AnnualRate = ApplRates(ValReturn.NumericValue - 1).Value.NumericValue
                TotCost += UnitCost * AnnualUnits * AnnualRate
                i -= 1
                j += 1
                ValReturn = Nothing

            Loop

            Return TotCost

        End Function

    End Class
#End Region ' #PMCostImplementation
#Region "#UnprocessedableForumula"

    Class UnknownFunctionVisitor
        Inherits DevExpress.Spreadsheet.Formulas.ExpressionVisitor
        Public HasUnknownFunction As Boolean
        Public Overrides Sub Visit(ByVal TestExpression As DevExpress.Spreadsheet.Formulas.UnknownFunctionExpression)
            MyBase.Visit(TestExpression)
            HasUnknownFunction = True

        End Sub

    End Class

#End Region

End Namespace
