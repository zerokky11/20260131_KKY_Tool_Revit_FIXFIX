Option Strict On
Option Explicit On

Imports System
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports NPOI.SS.UserModel

Namespace Global.KKY_Tool_Revit.Infrastructure
    Public Module NpoiSheetAutoSizeCompat
        <Extension()>
        Public Sub TrackAllColumnsForAutoSizing(sheet As ISheet)
            If sheet Is Nothing Then Return
            Try
                Dim mi As MethodInfo =
                    sheet.GetType().GetMethod("TrackAllColumnsForAutoSizing", BindingFlags.Instance Or BindingFlags.Public)
                If mi IsNot Nothing AndAlso mi.GetParameters().Length = 0 Then
                    mi.Invoke(sheet, Nothing)
                End If
            Catch
                ' ignore (hint only)
            End Try
        End Sub
    End Module
End Namespace
