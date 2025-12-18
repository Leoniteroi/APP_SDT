Option Strict On
Option Explicit On

Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices

Namespace SDT.Drawing

    Public NotInheritable Class LayerService
        Private Sub New()
        End Sub

        Public Shared Sub Ensure(tr As Transaction, db As Database, layerName As String, Optional colorIndex As Short = 1)
            If String.IsNullOrWhiteSpace(layerName) Then Exit Sub

            Dim lt As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            If lt.Has(layerName) Then Exit Sub

            lt.UpgradeOpen()

            Dim ltr As New LayerTableRecord() With {
                .Name = layerName,
                .Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex)
            }

            lt.Add(ltr)
            tr.AddNewlyCreatedDBObject(ltr, True)
        End Sub
    End Class

End Namespace
