Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Imports Autodesk.Civil.DatabaseServices

Namespace SDT.Civil.Commands

    Public Class LabelAssembly

        <CommandMethod("SDT_LABEL_ASSEMBLY_NAME")>
        Public Sub LabelAssemblyName_Field_MText()

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim pso As New PromptSelectionOptions() With {
                .MessageForAdding = Environment.NewLine & "Selecione os Assemblies (AeccDbAssembly): "
            }

            Dim psr As PromptSelectionResult = ed.GetSelection(pso)
            If psr.Status <> PromptStatus.OK OrElse psr.Value Is Nothing Then Return

            Dim createdTextIds As New List(Of ObjectId)()

            ' ---------------- CRIA OS MTEXTS COM FIELD ----------------
            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim ms As BlockTableRecord =
                    CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                For Each selObj As SelectedObject In psr.Value
                    If selObj Is Nothing OrElse selObj.ObjectId.IsNull Then Continue For

                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity =
                        TryCast(tr.GetObject(selObj.ObjectId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
                    If acEnt Is Nothing Then Continue For

                    Dim asm As Autodesk.Civil.DatabaseServices.Assembly =
                        TryCast(acEnt, Autodesk.Civil.DatabaseServices.Assembly)
                    If asm Is Nothing Then Continue For

                    Dim ext As Extents3d = acEnt.GeometricExtents
                    Dim xCenter As Double = (ext.MinPoint.X + ext.MaxPoint.X) / 2.0
                    Dim yTop As Double = ext.MaxPoint.Y
                    Dim insPt As New Point3d(xCenter, yTop + 2.0, 0.0)

                    Dim objIdDec As Long = asm.ObjectId.OldIdPtr.ToInt64()
                    Dim c14 As Char = Microsoft.VisualBasic.Strings.ChrW(14)

                    Dim fieldCode As String =
                        "%<\AcObjProp Object(%<\_ObjId " & objIdDec & ">%).Name \f """ &
                        c14 & "%tc1" & c14 & """>%"

                    Dim mt As New MText()
                    mt.SetDatabaseDefaults()
                    mt.Location = insPt
                    mt.TextHeight = 0.5
                    mt.Attachment = AttachmentPoint.BottomCenter
                    mt.Contents = fieldCode

                    ms.AppendEntity(mt)
                    tr.AddNewlyCreatedDBObject(mt, True)

                    createdTextIds.Add(mt.ObjectId)
                Next

                tr.Commit()
            End Using

            If createdTextIds.Count = 0 Then Return

            ' Seleciona os textos criados
            ed.SetImpliedSelection(createdTextIds.ToArray())

            ' Atualiza SOMENTE os fields contidos nesses MTexts (sem REGENALL)
            UpdateOnlyFields(db, ed, createdTextIds)

            ' Regen leve
            ed.Regen()

        End Sub

        ' ===================== HELPERS (DENTRO DA CLASS) =====================

        ''' <summary>
        ''' Atualiza somente os Fields embutidos nos MTexts fornecidos, varrendo o ExtensionDictionary.
        ''' </summary>
        Private Shared Sub UpdateOnlyFields(db As Database, ed As Editor, ids As IEnumerable(Of ObjectId))

            Using tr As Transaction = db.TransactionManager.StartTransaction()

                For Each id As ObjectId In ids
                    If id.IsNull Then Continue For

                    Dim mt As MText = TryCast(tr.GetObject(id, OpenMode.ForWrite), MText)
                    If mt Is Nothing Then Continue For
                    If Not mt.HasFields Then Continue For
                    If mt.ExtensionDictionary.IsNull Then Continue For

                    Dim extDict As DBDictionary =
                        CType(tr.GetObject(mt.ExtensionDictionary, OpenMode.ForRead), DBDictionary)

                    EvaluateFieldsInDictionary(tr, extDict, db)

                    mt.RecordGraphicsModified(True)
                Next

                tr.Commit()
            End Using

        End Sub

        Private Shared Sub EvaluateFieldsInDictionary(tr As Transaction, dict As DBDictionary, db As Database)

            For Each entry As DBDictionaryEntry In dict

                Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject =
    TryCast(tr.GetObject(entry.Value, OpenMode.ForWrite), Autodesk.AutoCAD.DatabaseServices.DBObject)
                If obj Is Nothing Then Continue For

                Dim f As Field = TryCast(obj, Field)
                If f IsNot Nothing Then
                    f.Evaluate(FieldEvaluationOptions.Automatic, db)
                    Continue For
                End If

                Dim subDict As DBDictionary = TryCast(obj, DBDictionary)
                If subDict IsNot Nothing Then
                    EvaluateFieldsInDictionary(tr, subDict, db)
                End If

            Next

        End Sub

    End Class

End Namespace