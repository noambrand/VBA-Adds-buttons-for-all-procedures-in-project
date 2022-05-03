Attribute VB_Name = "ProceduresList"
Option Explicit
' ----------------------------------------------------------------
' Purpose: called by RunCreateButton and Lists All VBA Procedures
' ----------------------------------------------------------------
'' set a reference to the Microsoft Visual Basic for Applications Extensibility 5.3 Library
Function ProceduresListPrint() As Variant
  ' Declare Excel workbook access variables.
  Dim App As Excel.Application
  Dim sOutput() As String
  ' Declare workbook macros access variables.
  Dim VBProj As VBIDE.VBProject
  Dim VBComp As VBIDE.VBComponent
  Dim vbMod As VBIDE.CodeModule
  ' Declare miscellaneous variables.
  Dim iRow As Long
  Dim iLine As Long
  Dim sProcname As String
  Dim ModuleName As String
  Dim pk As vbext_ProcKind
  '''''''
  ''''https://docs.microsoft.com/en-us/office/vba/api/access.module.procofline
   pk = vbext_pk_Proc
  ''''''''''
  Set App = Excel.Application
For Each VBProj In App.VBE.VBProjects
iRow = 0
    ' check for protected project
    On Error Resume Next
    Set VBComp = VBProj.VBComponents(1)
    On Error GoTo 0
    If Not VBComp Is Nothing Then
      ' Iterate through each component in the project.
            For Each VBComp In VBProj.VBComponents
                  ModuleName = VBComp.name
                  ' Find the code module for the project.
                  Set vbMod = VBComp.CodeModule
                  ' Scan through the code module, looking for procedures.
                  iLine = 1
                  Do While iLine < vbMod.CountOfLines
                          sProcname = vbMod.ProcOfLine(iLine, pk)
                          If sProcname <> vbNullString Then
                              iRow = iRow + 1
                              ReDim Preserve sOutput(1 To iRow)
                              sOutput(iRow) = sProcname
                              iLine = iLine + vbMod.ProcCountLines(sProcname, pk)
                          Else
                            ' This line has no procedure, so go to the next line.
                            iLine = iLine + 1
                          End If
                  Loop
                  ' clean up
                  Set vbMod = Nothing
                  Set VBComp = Nothing
            Next
    Else
        ReDim Preserve sOutput(1 To 3)
        sOutput(3) = "Project protected"
    End If

    If UBound(sOutput) = 2 Then
      ReDim Preserve sOutput(1 To 3)
      sOutput(3) = "No code in project"
    End If

    Set VBProj = Nothing
Next

ProceduresListPrint = sOutput()
End Function

