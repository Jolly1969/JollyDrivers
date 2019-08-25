

Imports System.Data.SqlClient
Imports System.Web.Management
Public Class roles_remove

    Inherits ClassMaster

    Dim rolesArray() As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ' Bind roles to ListBox.
            rolesArray = Roles.GetAllRoles()
            ListBox_AllRoles.DataSource = rolesArray
            ListBox_AllRoles.DataBind()
        End If
    End Sub


    Protected Sub Button_RemoveRoll_Click(sender As Object, e As EventArgs) Handles Button_RemoveRoll.Click
        If MsgBox("This will remove all users from this role and then remove the role itself! Continue?", vbYesNo, "Warning!") = vbYes Then
            Dim RemoveRollName As String = Convert.ToString(ListBox_AllRoles.SelectedItem)
            Dim SQL As String = "SELECT aspnet_Users.UserName" &
                            " FROM aspnet_Roles INNER JOIN" &
                            " aspnet_UsersInRoles ON aspnet_Roles.RoleId = aspnet_UsersInRoles.RoleId INNER JOIN" &
                            " aspnet_Users ON aspnet_UsersInRoles.UserId = aspnet_Users.UserId" &
                            " WHERE (aspnet_Roles.RoleName = @RoleName)"
            Dim conn As New SqlConnection(ConnString_SqlServices)
            Dim command As New SqlCommand(SQL, conn)
            command.Parameters.Add("@RoleName", SqlDbType.VarChar).Value = RemoveRollName
            conn.Open()
            Dim drR As SqlDataReader = command.ExecuteReader
            command = Nothing
            Do While drR.Read
                Dim UserName As String = Convert.ToString(drR.Item("UserName"))
                Roles.RemoveUserFromRole(UserName, RemoveRollName)
            Loop
            drR.Close()
            conn.Close()
            conn = Nothing
            Roles.DeleteRole(RemoveRollName)
            rolesArray = Roles.GetAllRoles()
            ListBox_AllRoles.DataSource = rolesArray
            ListBox_AllRoles.DataBind()
        End If

    End Sub
    'Protected Sub populateAllUsersInRole()
    '    ListBox_UsersInRole.Items.Clear()
    '    Dim RollName As String = Convert.ToString(ListBox_AllRoles.SelectedItem)
    '    Dim SQL As String = "SELECT aspnet_Users.UserName" &
    '                        " FROM aspnet_Roles INNER JOIN" &
    '                        " aspnet_UsersInRoles ON aspnet_Roles.RoleId = aspnet_UsersInRoles.RoleId INNER JOIN" &
    '                        " aspnet_Users ON aspnet_UsersInRoles.UserId = aspnet_Users.UserId" &
    '                        " WHERE (aspnet_Roles.RoleName = @RoleName)"
    '    Dim conn As New SqlConnection(ConnString_SqlServices)
    '    Dim command As New SqlCommand(SQL, conn)
    '    command.Parameters.Add("@RoleName", SqlDbType.VarChar).Value = RollName
    '    conn.Open()
    '    Dim drR As SqlDataReader = command.ExecuteReader
    '    command = Nothing
    '    Do While drR.Read
    '        Dim UserName As String = Convert.ToString(drR.Item("UserName"))
    '        ListBox_UsersInRole.Items.Add(UserName)
    '    Loop
    '    drR.Close()
    '    conn.Close()
    '    conn = Nothing
    'End Sub

    'Protected Sub ListBox_AllRoles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox_AllRoles.SelectedIndexChanged
    '    Call populateAllUsersInRole()
    '    'Call populateAllUsersNotInRole()
    'End Sub

    'Protected Sub Button_RemoveUserFromRole_Click(sender As Object, e As EventArgs) Handles Button_RemoveUserFromRole.Click
    '    Dim RemoveUser As String
    '    RemoveUser = ListBox_UsersInRole.SelectedValue
    '    Dim RemoveFromRole As String = Convert.ToString(ListBox_AllRoles.SelectedItem)
    '    Roles.RemoveUserFromRole(RemoveUser, RemoveFromRole)
    '    Call populateAllUsersInRole()
    'End Sub

    'Protected Sub populateAllUsersNotInRole()
    '    ListBox_Allusers.Items.Clear()
    '    Dim SQL As String = "SELECT UserName FROM aspnet_Users ORDER BY UserName"
    '    Dim conn As New SqlConnection(ConnString_SqlServices)
    '    Dim command As New SqlCommand(SQL, conn)
    '    conn.Open()
    '    Dim drR As SqlDataReader = command.ExecuteReader
    '    command = Nothing
    '    Do While drR.Read
    '        Dim UserName As String = Convert.ToString(drR.Item("UserName"))
    '        Dim listitem As New ListItem(UserName)
    '        If Not ListBox_UsersInRole.Items.Contains(listitem) Then
    '            ListBox_Allusers.Items.Add(listitem)
    '        End If
    '    Loop
    '    drR.Close()
    '    conn.Close()
    '    conn = Nothing
    'End Sub
End Class