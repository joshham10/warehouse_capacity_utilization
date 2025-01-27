# Warehouse Capacity Utilization Project
Used solver to minimize cost of warehouses in first scenario by improving capacity utilization. Automated process for 100 scenarios with macro recorder and VBA. Displayed findings in tableau to analyze trends.

## Project Steps

### 1. Excel and Solver
Determined how to minimize cost with the use of solver tool in excel for the first scenario. Changed the quantity of Warehouses A-C and added constraints to the supply of each warehouse and the demand for every store. 

### 2. Macro Recorder and VBA

- **Record Solver**: Record macro while using solver to automatically write code

- **Edit in VBA to Automate**: Changed code to automatically run macro through all 100 scenarios

```vba
    
    Sub Solver()
    ' Loop through all "Scenario_1" to "Scenario_100"
    Dim ws As Worksheet
    Dim scenarioName As String
    Dim i As Long
    
    For i = 1 To 100
        scenarioName = "Scenario_" & i
        
        ' Attempt to set and activate the worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(scenarioName)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' Activate the worksheet to ensure context
            ws.Activate
            
            ' Clear any previous Solver settings
            Application.Run "SolverReset"
            
            ' Perform Solver operations on the active sheet
            With ws
                ' Set default starting values for B10:E12
                .Range("B10:E12").Value = 1
                
                ' Set formulas
                .Range("F10").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
                .Range("F10:F12").FillDown
                .Range("B13").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
                .Range("B13:E13").FillRight
                .Range("B14").FormulaR1C1 = "=SUMPRODUCT(R[-12]C:R[-10]C[3], R[-4]C:R[-2]C[3])"
                
                ' Add Solver constraints using positional arguments
                Application.Run "SolverAdd", "F10:F12", 1, "F2:F4"
                Application.Run "SolverAdd", "B13:E13", 3, "B5:E5"
                
                ' Set Solver objective using positional arguments
                Application.Run "SolverOk", "B14", 2, 0, "B10:E12", 2, "Simplex LP"
                
                ' Run Solver without showing dialog
                Dim solverResult As Long
                solverResult = Application.Run("SolverSolve", True)
                
                ' Check if Solver found a solution
                If solverResult <> 0 Then
                    MsgBox "Solver could not find a solution for " & scenarioName, vbExclamation
                End If
                
                ' Finalize Solver
                Application.Run "SolverFinish", 1
            End With
        Else
            MsgBox "Worksheet " & scenarioName & " does not exist.", vbExclamation
        End If
        
        ' Reset the ws object for the next iteration
        Set ws = Nothing
    Next i
    
    ' Notify user when all scenarios are processed
    MsgBox "Solver automation completed for all scenarios.", vbInformation
    End Sub
```

- **Add Output Sheet**: Add new Sub to publish findings to a separate output sheet

```vba

Sub GenerateOutput()
    Dim wsOutput As Worksheet
    Dim wsScenario As Worksheet
    Dim scenarioName As String
    Dim totalCost As Double
    Dim i As Long
    Dim outputRow As Long
    
    ' Set the Output sheet
    Set wsOutput = ThisWorkbook.Sheets("Output")
    
    ' Clear existing data in the Output sheet
    wsOutput.Cells.Clear
    
    ' Write headers in the Output sheet
    wsOutput.Cells(1, 1).Value = "Scenario"
    wsOutput.Cells(1, 2).Value = "Total Cost"
    
    ' Start writing data from the second row
    outputRow = 2
    
    ' Loop through all "Scenario_1" to "Scenario_100"
    For i = 1 To 100
        scenarioName = "Scenario_" & i
        
        ' Check if the scenario sheet exists
        On Error Resume Next
        Set wsScenario = ThisWorkbook.Sheets(scenarioName)
        On Error GoTo 0
        
        If Not wsScenario Is Nothing Then
            ' Get the total cost from cell B14 of the scenario sheet
            totalCost = wsScenario.Range("B14").Value
            
            ' Write the scenario name and total cost to the Output sheet
            wsOutput.Cells(outputRow, 1).Value = scenarioName
            wsOutput.Cells(outputRow, 2).Value = totalCost
            
            ' Move to the next row in the Output sheet
            outputRow = outputRow + 1
        End If
    Next i
    
    ' Notify the user
    MsgBox "Output table generated successfully.", vbInformation
    End Sub
```

### 3. Tableau
- **Total Cost**: Used basic bar chart to display cost for every scenario
- **Warehouse Capacity Utilization**: Used horizontal bar chart using warehouse averages from each scenario
- **Warehouse Relationship to Stores**: Created separtate bar charts for each store aiming to find which warehouse was utilized by most stores
- **Transportation Costs and Cost Variance**: Looking to find relationship between cost variance and transportation costs using scatter plot
- **Demand and Cost**: Looking to see if total demand and total transportation cost has a positive or negative correlation

## Findings
-**Warehouse Capacity Utilization**: Warehouse 1 utilized their capacity the best with warehouse 2 being the second most effective and warehouse 3 being the least effective

-**Warehouse Relationship to Stores**: Warehouse 2 was likely the most important to the needs of store 1, warehouse 3 serves the most importnat role in the majority of the scenarios and warehouse seems to have the least amount of impact on the scenarios

-**Transportation and Cost Variance**: As transportation costs increases, so does cost variance. This is likely due to an element of variability and change, requiring more investment in flexible but costly transportation

-**Impact of Demand on Cost**: There is a positive correlation between both the variables meaning that as demand increases, total transportation costs also rise likely because of the increased shipping quantities resulting in further cost

## Author - Joshua Hampton

The project is part of my portfolio, displaying my skills in VBA, Excel, Tableau, and my understanding of the visulizations from the data, all important skills for a business/data analyst role. If you have any questions or thoughts please feel free to contact me!


Contact Information:
Email - josh.hampton735@gmail.com
LinkedIn: https://www.linkedin.com/in/josh-hampton/
