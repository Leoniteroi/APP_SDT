Option Strict On
Option Explicit On

Imports System
Imports Autodesk.AutoCAD.DatabaseServices

Namespace SDT.Core

    Public NotInheritable Class TransactionRunner
        Private Sub New()
        End Sub

        Public Shared Sub RunWrite(db As Database, action As Action(Of Transaction))
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            If action Is Nothing Then Throw New ArgumentNullException(NameOf(action))

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                action(tr)
                tr.Commit()
            End Using
        End Sub

        Public Shared Function RunWrite(Of T)(db As Database, func As Func(Of Transaction, T)) As T
            If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
            If func Is Nothing Then Throw New ArgumentNullException(NameOf(func))

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim result As T = func(tr)
                tr.Commit()
                Return result
            End Using
        End Function
    End Class

End Namespace
