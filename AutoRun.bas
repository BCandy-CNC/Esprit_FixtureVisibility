Attribute VB_Name = "AutoRun"
Option Explicit

Dim SimSup As cls_SimulationSuppression

Sub AutoOpen()
    Set SimSup = New cls_SimulationSuppression
End Sub
