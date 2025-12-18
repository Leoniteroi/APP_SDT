Option Strict On
Option Explicit On

Imports Autodesk.AutoCAD.DatabaseServices
Imports SDT.Drawing

Public NotInheritable Class LayerUtils
    Private Sub New()
    End Sub

    Public Shared Sub Ensure(tr As Transaction, db As Database, layerName As String, Optional colorIndex As Short = 1)
        LayerService.Ensure(tr, db, layerName, colorIndex)
    End Sub
End Class
