Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Namespace SDT.Civil

    Public NotInheritable Class CorridorSurfaceBuilderService

        Public Class SurfaceSpec
            Public Property Suffix As String
            Public Property IniFile As String
            Public Property DefaultCodes As String()
            Public Property OverhangMode As String
        End Class

        Private Sub New()
        End Sub



        Private Shared Function TryGetCorridorSurface(corr As Corridor, surfaceName As String) As CorridorSurface
            Try
                For Each cs As CorridorSurface In corr.CorridorSurfaces
                    If cs IsNot Nothing AndAlso cs.Name.Equals(surfaceName) Then
                        Return cs
                    End If
                Next
            Catch
            End Try
            Return Nothing
        End Function

        Public Shared Sub BuildAll(tr As Transaction,
                                   civDoc As CivilDocument,
                                   ed As Editor,
                                   spec As SurfaceSpec,
                                   Optional recreateSurface As Boolean = True)

            Dim codes As List(Of String) = IniCodesRepository.LoadDefaultCodes(spec.IniFile, spec.DefaultCodes)

            For Each corrId As ObjectId In civDoc.CorridorCollection
                Dim corr As Corridor = TryCast(tr.GetObject(corrId, OpenMode.ForWrite), Corridor)
                If corr Is Nothing Then Continue For

                Dim surfName As String = corr.Name & spec.Suffix


                'Tenta pegar a surface existente
                Dim corSurf As CorridorSurface = TryGetCorridorSurface(corr, surfName)

                If corSurf IsNot Nothing Then
                    If recreateSurface Then
                        'Remove e recria (comportamento atual)
                        RemoveCorridorSurfaceIfExists(corr, surfName)
                        'corSurf = Nothing
                        Try
                            corSurf = corr.CorridorSurfaces.Add(surfName)
                        Catch ex As Exception
                            ed.WriteMessage(Environment.NewLine & "[SDT] Falha criando CorridorSurface '" & surfName & "': " & ex.Message)
                            Continue For
                        End Try
                    Else
                        'NÃO deleta: só atualiza os codes e segue para o próximo corredor
                        'UpdateLinkCodes(corSurf, codes, ed)

                        'Opcional: aplicar overhang mesmo sem recriar
                        If Not String.IsNullOrWhiteSpace(spec.OverhangMode) Then
                            OverhangCorrector.TryApply(corSurf, spec.OverhangMode, ed)
                        End If

                        'Continue For
                    End If
                Else
                    Try
                        corSurf = corr.CorridorSurfaces.Add(surfName)
                    Catch ex As Exception
                        ed.WriteMessage(Environment.NewLine & "[SDT] Falha criando CorridorSurface '" & surfName & "': " & ex.Message)
                        Continue For
                    End Try
                End If

                AddLinkCodes(corSurf, codes, ed)

                Try
                    corSurf.Boundaries.AddCorridorExtentsBoundary(surfName & "_OUTER")
                Catch
                End Try

                If Not String.IsNullOrWhiteSpace(spec.OverhangMode) Then
                    OverhangCorrector.TryApply(corSurf, spec.OverhangMode, ed)
                End If

                corr.Rebuild()
            Next
        End Sub


        Private Shared Sub RemoveCorridorSurfaceIfExists(corr As Corridor, surfaceName As String)
            Try
                For Each cs As CorridorSurface In corr.CorridorSurfaces
                    If cs IsNot Nothing AndAlso cs.Name.Equals(surfaceName) Then ', StringComparison.OrdinalIgnoreCase) Then
                        Try
                            corr.CorridorSurfaces.Remove(cs)
                        Catch
                        End Try
                        Exit For
                    End If
                Next
            Catch
            End Try
        End Sub

        Private Shared Sub AddLinkCodes(corSurf As CorridorSurface, codes As List(Of String), ed As Editor)
            For Each code As String In codes
                Try
                    corSurf.AddLinkCode(code, True)
                Catch ex As ArgumentException
                    ed.WriteMessage(Environment.NewLine & "[SDT] LinkCode inválido '" & code & "' em '" & corSurf.Name & "': " & ex.Message)
                Catch
                End Try
            Next
        End Sub

    End Class

End Namespace
