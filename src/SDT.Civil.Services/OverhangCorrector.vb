Option Strict On
Option Explicit On

Imports System
Imports System.Reflection
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.DatabaseServices

Namespace SDT.Civil

    Public NotInheritable Class OverhangCorrector
        Private Sub New()
        End Sub

        Public Shared Function TryApply(corSurf As CorridorSurface, mode As String, ed As Editor) As Boolean
            If corSurf Is Nothing Then Return False
            If String.IsNullOrWhiteSpace(mode) Then Return False

            Dim nl As String = Environment.NewLine

            Try
                Dim t As Type = corSurf.GetType()

                ' 1) Método ApplyOverhangCorrection(...)
                Dim mi As MethodInfo =
                    t.GetMethod("ApplyOverhangCorrection", BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)

                If mi IsNot Nothing Then
                    Dim pars As ParameterInfo() = mi.GetParameters()
                    If pars.Length = 1 Then
                        Dim pType As Type = pars(0).ParameterType

                        ' 1a) String
                        If pType Is GetType(String) Then
                            mi.Invoke(corSurf, New Object() {mode})
                            ed.WriteMessage(nl & "[SDT] Overhang aplicado por método (String): " & mode)
                            TryRebuildSurface(corSurf, ed)
                            Return True
                        End If

                        ' 1b) Enum
                        If pType.IsEnum Then
                            Dim enumVal As Object = ParseEnumValue(pType, mode)
                            mi.Invoke(corSurf, New Object() {enumVal})
                            ed.WriteMessage(nl & "[SDT] Overhang aplicado por método (Enum " & pType.Name & "): " & enumVal.ToString())
                            TryRebuildSurface(corSurf, ed)
                            Return True
                        End If
                    End If
                End If

                ' 2) Propriedades (varia por versão)
                Dim candidates As String() = New String() {"OverhangCorrection", "OverhangCorrectionType", "OverhangMode"}

                For Each propName As String In candidates
                    Dim pi As PropertyInfo =
                        t.GetProperty(propName, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)

                    If pi Is Nothing OrElse Not pi.CanWrite Then Continue For

                    Dim pType As Type = pi.PropertyType

                    If pType Is GetType(String) Then
                        pi.SetValue(corSurf, mode, Nothing)
                        ed.WriteMessage(nl & "[SDT] Overhang aplicado por propriedade (" & propName & "=String): " & mode)
                        TryRebuildSurface(corSurf, ed)
                        Return True
                    End If

                    If pType.IsEnum Then
                        Dim enumVal As Object = ParseEnumValue(pType, mode)
                        pi.SetValue(corSurf, enumVal, Nothing)
                        ed.WriteMessage(nl & "[SDT] Overhang aplicado por propriedade (" & propName & "=Enum " & pType.Name & "): " & enumVal.ToString())
                        TryRebuildSurface(corSurf, ed)
                        Return True
                    End If

                    If pType Is GetType(Integer) Then
                        Dim intVal As Integer = TryParseInt(mode)
                        pi.SetValue(corSurf, intVal, Nothing)
                        ed.WriteMessage(nl & "[SDT] Overhang aplicado por propriedade (" & propName & "=Integer): " & intVal.ToString())
                        TryRebuildSurface(corSurf, ed)
                        Return True
                    End If
                Next

                ' 3) Diagnóstico
                ed.WriteMessage(nl & "[SDT] Overhang: não encontrei método/propriedade compatível. Surface='" &
                               corSurf.Name & "', Mode='" & mode & "'.")
                DumpOverhangMembers(corSurf, ed)
                Return False

            Catch ex As Exception
                ed.WriteMessage(nl & "[SDT] OverhangCorrection falhou em '" & corSurf.Name & "': " & ex.Message)
                Return False
            End Try
        End Function

        Private Shared Sub TryRebuildSurface(corSurf As CorridorSurface, ed As Editor)
            Try
                Dim nl As String = Environment.NewLine
                Dim t As Type = corSurf.GetType()
                Dim mi As MethodInfo =
                    t.GetMethod("Rebuild", BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)

                If mi IsNot Nothing Then
                    Dim pars As ParameterInfo() = mi.GetParameters()
                    If pars.Length = 0 Then
                        mi.Invoke(corSurf, Nothing)
                        ed.WriteMessage(nl & "[SDT] Rebuild() da CorridorSurface executado.")
                    End If
                End If
            Catch
                ' best effort
            End Try
        End Sub

        Private Shared Function ParseEnumValue(enumType As Type, mode As String) As Object
            Dim raw As String = mode.Trim()

            ' Parse direto
            Try
                Return [Enum].Parse(enumType, raw, True)
            Catch
                ' normaliza
                Dim normalized As String = raw.Replace(" ", "").Replace("-", "").Replace("_", "")
                Dim names As String() = [Enum].GetNames(enumType)

                For Each n As String In names
                    Dim nn As String = n.Replace(" ", "").Replace("-", "").Replace("_", "")
                    If String.Equals(nn, normalized, StringComparison.OrdinalIgnoreCase) Then
                        Return [Enum].Parse(enumType, n, True)
                    End If
                Next
            End Try

            ' fallback: primeiro valor do enum
            Dim values As Array = [Enum].GetValues(enumType)
            Return values.GetValue(0)
        End Function

        Private Shared Function TryParseInt(s As String) As Integer
            Dim v As Integer
            If Integer.TryParse(s.Trim(), v) Then Return v
            Return 0
        End Function

        Private Shared Sub DumpOverhangMembers(corSurf As CorridorSurface, ed As Editor)
            Try
                Dim nl As String = Environment.NewLine
                Dim t As Type = corSurf.GetType()

                ed.WriteMessage(nl & "[SDT] Dump members relacionados a Overhang/Correction:")

                Dim members As MemberInfo() = t.GetMembers(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                For Each m As MemberInfo In members
                    Dim name As String = m.Name
                    If name.IndexOf("Overhang", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       name.IndexOf("Correction", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        ed.WriteMessage(nl & "  - " & m.MemberType.ToString() & ": " & name)
                    End If
                Next

            Catch
                ' best effort
            End Try
        End Sub

    End Class

End Namespace
