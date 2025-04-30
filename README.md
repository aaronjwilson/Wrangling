### Wrangling or i didn't know i had data stack capabilities.
Musings on Data munging and the development of ETL pipelines from biological data sources.

life in a biological lab is more than just the data and the insights developed as there are navigations that go unspoken about the data: the quality, the management and the personalities that control it. the stochasticity of the 'navigations' require advanced maths to even begin to address the complexity. 


some thoughts on the process of getting data into a format that works best for getting the most information (graphical output, EDA, statistical measures)

```mermaid
stateDiagram
    [*] --> Fit_data_to_software
   

    Fit_data_to_software --> Fit_software_to_data
    Fit_software_to_data --> Fit_data_to_software
  
   Fit_software_to_data  --> [*]
```
For this part i will focus on developments in excel vbscript but future discussions will follow the development of bespoke methods using alternative scripting/programming languages such as: 

R/Python<br>
Macros made in javascript for googlesheets or O365.

###VBSCIPTING<br>
when thinking about the tools that i use on the regular to bring data into a format that allows the easiest and most efficient method of getting the most amount of detail. timing is of the ssesnce. Excel is a great program. extended with visual basic is even better. <br>


<p align="center">
<img title="Raw Data from file" alt="Alt text" src="raw.png" width="300" height="100" align="center">
</p>
<p align="center">  
<img title="Categorical Data from file" alt="Alt text" src="categorical.png" width="300" height="300" align="center">
</p>
<p align="center">
<img title="Removed Data from file" alt="Alt text" src="adjusted.png" width="300" height="100" align="center">
</p>


want create vbscript to combine the raw and categorical and output a database ready structure like this: 
<p align="center">
<img title="Transformation" alt="Alt text" src="transform.png" width="300" height="100" align="center">
</p>



To begin the work the structure of the data must be analyzed and either established or created. VBscript for an excel macro to interact with the elisa data. Started with removal of empty rows. Nonlinear program with the function created in a subroutine pitched well below the machination work. 
```vbscript
Sub Gen5Row()

' loopish Macro
Range("A1:A1700").Select
Application.Run "DelEmptyRow"
```

Next define a starting point and let the voodoo happen on loop. 
```vbscript
Range("B3").Select
Do Until IsEmpty(ActiveCell.Offset(22, -1))
    ActiveCell.Range("A1,A4,A7,A10,A13,A16,A19,A22").Select
    Selection.Copy
    ActiveCell.Offset(22, -1).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 1).Range("A1:L1").Select 'active cell A25
    Application.CutCopyMode = False
    Selection.Cut
    ActiveCell.Offset(-1, 12).Range("A1").Select
    ActiveSheet.Paste
        For i = 1 To 3
            ActiveCell.Offset(3, -12).Range("A1:L1").Select
            Selection.Cut
            ActiveCell.Offset(-1, 12).Range("A1").Select
            ActiveSheet.Paste
        Next i
   ActiveCell.Offset(12, -12).Range("A1").Select
Loop
```
Another reference to the deletion of empty rows.
```vbscript
    Range("O1:O1500").Select
    Application.Run "DelEmptyRow"
```


Nomenclature built around direct naming and the name built into the sheet tab
```vbscript
Range("A1").EntireRow.Insert
Range("A:C").EntireColumn.Insert
Range("A1")="Protein"
Range("B1")="StudyID"
Range("C1")="Timepoint"
Range("D1")="Participantid"
Range("A2:A30") = Mid(ThisWorkbook.Name, 10, 15)
Range("B2:B30") = Mid(ThisWorkbook.Name, 1, 8)
Range("C2:C30") = ActiveSheet.Name
End Sub
```


##LOOK HERE FOR THE SUBROUTINE USED TWICE ABOVE
```vbscript
'Attribute VB_Name = "DelEmptyRow"
Sub DelEmptyRow()
Rng = Selection.Rows.Count
ActiveCell.Offset(0, 0).Select
Application.ScreenUpdating = False
For i = 1 To Rng
If ActiveCell.Value = "" Then 
Selection.EntireRow.Delete
Else
ActiveCell.Offset(1, 0).Select
End If
Next i
Application.ScreenUpdating = True
End Sub
```
Sandbox to test out nomenclature subroutine.  Data ended up being so different between studies the manual method worked the best. 
```vbscript
Sub FileNomen()
'Range("A2") = Mid(ThisWorkbook.Name, 10, 15)
'Range("B2") = Mid(ThisWorkbook.Name, 1, 8)
'Range("C2") = ActiveSheet.Name
'Range("A2") = Mid(ThisWorkbook.Name, 12, 16)
'Range("B2") = Mid(ThisWorkbook.Name, 1, 10)

End Sub
```




links: 
https://upmath.me/


