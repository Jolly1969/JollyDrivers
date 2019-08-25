<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Administrator/admin.Master" CodeBehind="roles_create.aspx.vb" Inherits="JollyDrivers.roles_create" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <title>Create Role</title>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <form id="createRole-form" name="createRole-form" class="nobottommargin" action="#" method="post">
                <h3>Create Role</h3>

                <div class="col_full">
                    <label for="createRole-form-rolename">New role name:</label>
                    <asp:TextBox ID="RoleTextBox" runat="server" class="form-control not-dark"></asp:TextBox>
                </div>

                <div class="col_full">
                    <asp:Button Text="Create Role" ID="CreateRoleButton" runat="server" OnClick="CreateRole_OnClick" />
                </div>

                <div class="col_full">
                    <label for="createRole-form-ListBox_AllRoles">Existing Roles</label>
                    <asp:ListBox ID="ListBox_AllRoles" runat="server" Width="100%" Rows="10" AutoPostBack="False" Enabled="false"></asp:ListBox>
                </div>
            </form>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
