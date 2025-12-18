Option Strict On
Option Explicit On

Imports System
Imports System.Reflection
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Namespace SDT.Civil

    Public NotInheritable Class TinSurfaceService
        Private Sub New()
        End Sub

        Public Shared Function FindTinSurfaceIdByName(tr As Transaction,
                                                      civDoc As CivilDocument,
                                                      name As String) As ObjectId
            If tr Is Nothing OrElse civDoc Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then
                Return ObjectId.Null
            End If

            For Each id As ObjectId In civDoc.GetSurfaceIds()
                If id.IsNull Then Continue For
                Dim s As TinSurface = TryCast(tr.GetObject(id, OpenMode.ForRead), TinSurface)
                If s IsNot Nothing AndAlso String.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase) Then
                    Return id
                End If
            Next

            Return ObjectId.Null
        End Function

        Public Shared Sub SafeErase(tr As Transaction, id As ObjectId)
            If tr Is Nothing OrElse id.IsNull Then Exit Sub
            Try
                Dim dbo As Autodesk.AutoCAD.DatabaseServices.DBObject =
    TryCast(tr.GetObject(id, OpenMode.ForWrite, False), Autodesk.AutoCAD.DatabaseServices.DBObject)


                If dbo IsNot Nothing AndAlso Not dbo.IsErased Then dbo.Erase(True)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' CLONA uma TinSurface (cópia independente) usando DeepCloneObjects.
        ''' Substitui qualquer abordagem baseada em PasteSurface.
        ''' </summary>
        ''' <remarks>
        ''' - Não cria dependência/ligação entre superfícies.
        ''' - Mais estável para processamento em lote.
        ''' </remarks>
        Public Shared Function CloneTinSurface(db As Database,
                                                  tr As Transaction,
                                                  civDoc As CivilDocument,
                                                  sourceTinId As ObjectId,
                                                  newName As String,
                                                  ed As Editor) As ObjectId
            If db Is Nothing OrElse tr Is Nothing OrElse civDoc Is Nothing Then Return ObjectId.Null
            If sourceTinId.IsNull Then Return ObjectId.Null

            Dim src As TinSurface = TryCast(tr.GetObject(sourceTinId, OpenMode.ForRead), TinSurface)
            If src Is Nothing Then Return ObjectId.Null

            If Not String.IsNullOrWhiteSpace(newName) Then
                Dim existingId As ObjectId = FindTinSurfaceIdByName(tr, civDoc, newName)
                If Not existingId.IsNull Then SafeErase(tr, existingId)
            End If

            Dim ownerId As ObjectId = src.OwnerId
            If ownerId.IsNull Then Return ObjectId.Null

            Dim ids As New ObjectIdCollection()
            ids.Add(sourceTinId)

            Dim idMap As New IdMapping()
            db.DeepCloneObjects(ids, ownerId, idMap, False)

            Dim newId As ObjectId = ObjectId.Null
            For Each pair As IdPair In idMap
                If pair.Key = sourceTinId AndAlso pair.IsCloned Then
                    newId = pair.Value
                    Exit For
                End If
            Next

            If newId.IsNull Then Return ObjectId.Null

            If Not String.IsNullOrWhiteSpace(newName) Then
                Dim cloned As TinSurface = TryCast(tr.GetObject(newId, OpenMode.ForWrite), TinSurface)
                If cloned IsNot Nothing Then cloned.Name = newName
            End If

            Return newId
        End Function

        ''' <summary>
        ''' Compatibilidade: mantenha chamadas antigas.
        ''' Use <see cref="CloneTinSurface"/> no lugar.
        ''' </summary>
<Obsolete("Use CloneTinSurface no lugar de CopyFromTinSurface (evita PasteSurface e dependências).")>
Public Shared Function CopyFromTinSurface(db As Database,
                                          tr As Transaction,
                                          civDoc As CivilDocument,
                                          sourceTinId As ObjectId,
                                          newName As String,
                                          ed As Editor) As ObjectId
    Return CloneTinSurface(db, tr, civDoc, sourceTinId, newName, ed)
End Function


        Public Shared Sub RaiseTinSurface(tin As TinSurface, raise As Double)
            If tin Is Nothing Then Exit Sub
            If Math.Abs(raise) < 0.0000001 Then Exit Sub

            Try
                tin.RaiseSurface(raise)
                Exit Sub
            Catch
            End Try

            ' fallback reflection
            Try
                Dim t As Type = tin.GetType()
                Dim mi As MethodInfo =
                    t.GetMethod("RaiseSurface", BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                                Nothing, New Type() {GetType(Double)}, Nothing)

                If mi Is Nothing Then
                    mi = t.GetMethod("Raise", BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                                     Nothing, New Type() {GetType(Double)}, Nothing)
                End If

                If mi IsNot Nothing Then mi.Invoke(tin, New Object() {raise})
            Catch
            End Try
        End Sub

        Public Shared Sub ApplyOuterBoundary(tin As TinSurface,
                                     boundaryPolylineId As ObjectId,
                                     midOrdinate As Double,
                                     keepEntities As Boolean,
                                     ed As Editor)

            If tin Is Nothing Then Exit Sub
            If boundaryPolylineId.IsNull Then Exit Sub

            Try
                Dim bdObj As Object = tin.BoundariesDefinition
                If bdObj Is Nothing Then
                    If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & "[SDT] BoundariesDefinition = Nothing.")
                    Exit Sub
                End If

                Dim bdType As Type = bdObj.GetType()

                ' Procura um mtodo "Add*" que receba:
                '  - (ObjectId, Double, Enum boundaryType, Boolean)
                ' ou variaes com 3/4/5 params.
                Dim chosen As MethodInfo = Nothing
                Dim chosenParams As ParameterInfo() = Nothing

                For Each mi As MethodInfo In bdType.GetMethods(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                    Dim name As String = mi.Name
                    If name Is Nothing Then Continue For
                    If Not name.StartsWith("Add", StringComparison.OrdinalIgnoreCase) Then Continue For

                    Dim ps As ParameterInfo() = mi.GetParameters()
                    If ps Is Nothing Then Continue For
                    If ps.Length < 3 Then Continue For

                    ' Primeiro param ObjectId (ou ObjectIdCollection)
                    Dim p0 As Type = ps(0).ParameterType
                    Dim okP0 As Boolean =
                (p0 Is GetType(ObjectId)) OrElse (p0 Is GetType(ObjectIdCollection))

                    If Not okP0 Then Continue For

                    ' Segundo param costuma ser Double (mid-ordinate) OU algo similar
                    Dim okP1 As Boolean = (ps.Length >= 2 AndAlso ps(1).ParameterType Is GetType(Double))
                    If Not okP1 Then Continue For

                    ' Terceiro param costuma ser enum de boundary type
                    Dim p2 As Type = ps(2).ParameterType
                    Dim okP2 As Boolean = p2.IsEnum

                    If Not okP2 Then Continue For

                    ' Achou um candidato bom
                    chosen = mi
                    chosenParams = ps
                    Exit For
                Next

                If chosen Is Nothing Then
                    If ed IsNot Nothing Then
                        ed.WriteMessage(Environment.NewLine & "[SDT] No achei mtodo Add* compatvel em BoundariesDefinition: " & bdType.FullName)
                    End If
                    Exit Sub
                End If

                ' Monta argumentos conforme a assinatura encontrada
                Dim args(chosenParams.Length - 1) As Object

                ' arg0: ObjectId ou ObjectIdCollection
                If chosenParams(0).ParameterType Is GetType(ObjectId) Then
                    args(0) = boundaryPolylineId
                Else
                    Dim col As New ObjectIdCollection()
                    col.Add(boundaryPolylineId)
                    args(0) = col
                End If

                ' arg1: mid-ordinate
                args(1) = midOrdinate

                ' arg2: enum boundary type (Outer)
                Dim enumType As Type = chosenParams(2).ParameterType
                Dim outerValue As Object = [Enum].Parse(enumType, "Outer", True)
                args(2) = outerValue

                ' Demais args: tenta preencher boolean "keepEntities" quando existir
                For i As Integer = 3 To chosenParams.Length - 1
                    Dim pt As Type = chosenParams(i).ParameterType

                    If pt Is GetType(Boolean) Then
                        args(i) = keepEntities
                    ElseIf pt Is GetType(Double) Then
                        ' alguns overloads repetem tolerncia etc.
                        args(i) = 0.0R
                    ElseIf pt Is GetType(Integer) Then
                        args(i) = 0
                    Else
                        ' desconhecido -> Nothing
                        args(i) = Nothing
                    End If
                Next

                chosen.Invoke(bdObj, args)

            Catch ex As Exception
                If ed IsNot Nothing Then
                    ed.WriteMessage(Environment.NewLine & "[SDT] ApplyOuterBoundary falhou: " & ex.Message)
                End If
            End Try

        End Sub


    End Class

End Namespace
