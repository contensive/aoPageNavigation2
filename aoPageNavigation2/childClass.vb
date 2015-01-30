

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.Addon.PageNavigation2
    '
    ' Sample Vb addon
    '
    Public Class childClass
        Inherits AddonBaseClass
        '
        ' - Create a Contensive Addon record, set the dotnet class full name to yourNameSpaceName.yourClassName
        '
        '=====================================================================================
        ' addon api
        '=====================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String
            Try
                returnHtml = commonModule.GetPageNavigation(CP, commonModule.PageTypeChild, 2)
            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = "Error response"
            End Try
            Return returnHtml
        End Function
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in sampleClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
    End Class
End Namespace
