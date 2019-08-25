
Imports System.Net.Mail
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Threading

Imports System.Drawing
Imports System.Drawing.Imaging

Public Class ClassMaster
    Inherits System.Web.UI.Page

    Public ReadOnly Property ConnString_SqlServices() As String
        Get
            Return ConfigurationManager.ConnectionStrings("SqlServices").ConnectionString
        End Get


    End Property



End Class
