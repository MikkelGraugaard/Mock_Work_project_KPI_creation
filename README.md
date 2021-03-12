# Mock_Work_project_KPI_creation
Showcasing example of work project


### Functions used in excel sheet

- MaxIfs()
- Ifelse()
- Max()
- CountIf()
- Index(Match())
- IfError()
- Today()

Mainly using above functions to extraxt and create KPI's from exsting data. The excel functions are used both seperated and nested. Aswells as calculations are made between them. 



## VBA Cod.
Example of the VBA code that uses the KPI's and adittional data in the sheet to dictate what should happen with the IT-asset


*First running some code to optimize the excel sheet prior to using the function*
```
Public CalcState As Long

Public EventState As Boolean

Public PageBreakState As Boolean
```
 
 
 
```
Sub OptimizeCode_Begin()
_______________________________________________
 

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub
```


 
```
Sub OptimizeCode_End()
_________________________________________________
 

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.EnableEvents = EventState
Application.ScreenUpdating = True

End Sub
```
 

*Making the function that is to be used in the excel sheet*
```
Function Bortskaffes(CellRef1 As Date, CellRef2 As String, CellRef3 As Integer, CellRef4 As Integer, CellRef5 As String) As String
__________________________________________________________________________________
 
' First running the optimization code

Call OptimizeCode_Begin 


'Ser om computeren har nyere retirement end i dag, derefter tjekkese det om det er den nyeste computer personen har, hvis ja så ""

If CellRef1 > Date And CellRef3 = 1 Then
    Bortskaffes = "Nyeste PC og ikke overskredet retirement"

 

'Ser om personen har mere end 1 PC, hvis ja og det ikke er den nyeste så skal den tilbageleveres

ElseIf CellRef1 > Date And CellRef4 > 1 And CellRef3 = 0 Then
    Bortskaffes = "Mere end 1 PC og ikke nyeste PC"

   

'Ser som computeren har overskredet retirement date, hvis ja og personen har mere end 1 computer så bortskaffes
'Så tjekkes om det er den nyeste computer. hvilken der skal udskiftes

 

ElseIf CellRef1 <= Date And CellRef4 > 1 And CellRef3 = 1 Then
    Bortskaffes = "Overskredet retirement og nyeste PC"

ElseIf CellRef1 <= Date And CellRef3 = 0 Then
    Bortskaffes = "Overskredet retirement og ikke nyeste PC"

 

'Ser om retirement er udløbet og om personen kun har 1 computer, hvis ja så skal computeren udskiftes.

ElseIf CellRef1 <= Date And CellRef4 = 1 Then
    Bortskaffes = "Overskredet retirement og kun 1 PC"

 

'Ser om AD er false

ElseIf CellRef5 = "False" Then
    Bortskaffes = "Bortskaffes fordi AD er False"

   

ElseIf CellRef2 = "" Then
    Bortskaffes = "Mangler bruger"

Else: Bortskaffes = "Bortskaffes"

    End If

Call OptimizeCode_End

End Function
```
