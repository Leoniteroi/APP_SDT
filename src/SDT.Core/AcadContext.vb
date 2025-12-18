Option Strict On
Option Explicit On

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices

Namespace SDT.Core

    Public NotInheritable Class AcadContext
        Public ReadOnly Property Doc As Document
        Public ReadOnly Property Db As Database
        Public ReadOnly Property Ed As Editor
        Public ReadOnly Property CivDoc As CivilDocument

        Private Sub New(doc As Document, db As Database, ed As Editor, civDoc As CivilDocument)
            Me.Doc = doc
            Me.Db = db
            Me.Ed = ed
            Me.CivDoc = civDoc
        End Sub

        Public Shared Function TryCreate(ByRef ctx As AcadContext, ByRef err As String) As Boolean
            ctx = Nothing
            err = ""

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            If doc Is Nothing Then
                err = "Documento ativo não encontrado."
                Return False
            End If

            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim civDoc As CivilDocument = CivilApplication.ActiveDocument
            If civDoc Is Nothing Then
                err = "CivilDocument não disponível (abra um desenho do Civil 3D)."
                Return False
            End If

            ctx = New AcadContext(doc, db, ed, civDoc)
            Return True
        End Function
    End Class

End Namespace
