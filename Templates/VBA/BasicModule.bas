Attribute VB_Name = "FinanceUtilities"
'*******************************************************************************
' Finance Utilities Module
' Description: Template for creating financial utility functions in VBA
' Author: Finance Team
' Date: September 9, 2025
'*******************************************************************************

Option Explicit

'*******************************************************************************
' Constants
'*******************************************************************************
Public Const TAX_RATE As Double = 0.21 ' Example corporate tax rate

'*******************************************************************************
' Public Functions
'*******************************************************************************
Public Function CalculateNPV(rate As Double, cashFlows As Range) As Double
    ' Calculates Net Present Value for a series of cash flows
    ' Parameters:
    '   rate - Discount rate as a decimal (e.g., 0.08 for 8%)
    '   cashFlows - Range of cells containing the cash flow values
    
    Dim i As Long
    Dim npv As Double
    
    npv = 0
    
    For i = 1 To cashFlows.Count
        npv = npv + cashFlows(i).Value / ((1 + rate) ^ i)
    Next i
    
    CalculateNPV = npv
End Function

Public Function CalculateIRR(initialInvestment As Double, cashFlows As Range) As Double
    ' Simple wrapper for built-in IRR function
    ' Parameters:
    '   initialInvestment - The initial cash outflow (negative value)
    '   cashFlows - Range of cells containing the subsequent cash flow values
    
    ' Using built-in IRR function
    Dim fullCashFlowsArray() As Double
    Dim i As Long
    
    ReDim fullCashFlowsArray(0 To cashFlows.Count)
    
    fullCashFlowsArray(0) = initialInvestment
    
    For i = 1 To cashFlows.Count
        fullCashFlowsArray(i) = cashFlows(i).Value
    Next i
    
    CalculateIRR = IRR(fullCashFlowsArray)
End Function

'*******************************************************************************
' Private Helper Functions
'*******************************************************************************
Private Function FormatAsCurrency(value As Double, Optional currencySymbol As String = "$") As String
    ' Formats a value as currency with 2 decimal places
    ' Parameters:
    '   value - The numeric value to format
    '   currencySymbol - The currency symbol to use (default: $)
    
    FormatAsCurrency = currencySymbol & Format(value, "#,##0.00")
End Function
