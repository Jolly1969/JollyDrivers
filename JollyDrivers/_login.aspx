<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="_login.aspx.vb" Inherits="JollyDrivers._login" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">

    <meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <meta name="author" content="SemiColonWeb" />

    <!-- Stylesheets
	============================================= -->
    <link href="https://fonts.googleapis.com/css?family=Lato:300,400,400i,700|Raleway:300,400,500,600,700|Crete+Round:400i" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="css/bootstrap.css" type="text/css" />
    <link rel="stylesheet" href="style.css" type="text/css" />
    
    <link rel="stylesheet" href="css/dark.css" type="text/css" />
    <link rel="stylesheet" href="css/font-icons.css" type="text/css" />
    <link rel="stylesheet" href="css/animate.css" type="text/css" />
    <link rel="stylesheet" href="css/magnific-popup.css" type="text/css" />

    <link rel="stylesheet" href="css/responsive.css" type="text/css" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />

    <!-- Document Title
	============================================= -->
    <title>Login</title>

</head>

<body class="stretched">
    <form id="form2" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
        <!-- Document Wrapper
	============================================= -->
        <div id="wrapper" class="clearfix">

            <!-- Content
		============================================= -->
            <section id="content">

                <div class="content-wrap nopadding">

                    <div class="section nopadding nomargin" style="width: 100%; height: 100%; position: absolute; left: 0; top: 0; background: url('images/aerial-2178705_1920.jpg') center center no-repeat; background-size: cover;"></div>

                    <div class="section nobg full-screen nopadding nomargin">
                        <div class="container-fluid vertical-middle divcenter clearfix">

                           <%-- <div class="center">
                                <asp:Label ID="Label1" runat="server" Text="Jolly Drivers Collective"></asp:Label>
                                <a href="index.html"><img src="images/logo-dark.png" alt="Canvas Logo"></a>
                            </div>--%>

                            <div class="card divcenter noradius noborder" style="max-width: 400px; background-color: rgba(255,255,255,0.93);">
                                <div class="card-body" style="padding: 40px;">
                                    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                        <ContentTemplate>

                                            <asp:LoginView ID="LoginView1" runat="server">
                                                <AnonymousTemplate>
                                                    <asp:Login ID="Login1" runat="server">
                                                        <LayoutTemplate>


                                                            <form id="login-form" name="login-form" class="nobottommargin" action="#" method="post">
                                                                <h3>Login to your Account</h3>

                                                                <div class="col_full nobottommargin">
                                                                    <label for="login-form-username">Username:</label>
                                                                    <%--<input type="text" id="login-form-username" name="login-form-username" value="" class="form-control not-dark" />--%>
                                                                    <asp:TextBox ID="UserName" runat="server" class="form-control not-dark"></asp:TextBox>
                                                                                    <asp:RequiredFieldValidator ID="UserNameRequired" runat="server" ControlToValidate="UserName" ErrorMessage="User Name is required." ToolTip="User Name is required." ValidationGroup="ctl06$Login1">*</asp:RequiredFieldValidator>
                                                                </div>

                                                                <div class="col_full nobottommargin">
                                                                    <label for="login-form-password">Password:</label>
                                                                    <%--<input type="password" id="login-form-password" name="login-form-password" value="" class="form-control not-dark" />--%>
                                                                    <asp:TextBox ID="Password" runat="server" TextMode="Password" class="form-control not-dark"></asp:TextBox>
                                                                    <asp:RequiredFieldValidator ID="PasswordRequired" runat="server" ControlToValidate="Password" ErrorMessage="Password is required." ToolTip="Password is required." ValidationGroup="ctl06$Login1">*</asp:RequiredFieldValidator>
                                                                </div>

                                                                <div class="col_full">
                                                                    <asp:CheckBox ID="RememberMe" runat="server" Text="&nbsp;&nbsp;Remember me next time." /><br /><br />
                                                                    
                                                                    <asp:Button ID="LoginButton" runat="server" CommandName="Login" Text="Log In" ValidationGroup="ctl06$Login1" class="button button-3d button-black nomargin" />
                                                                    
                                                                    <%--<button class="button button-3d button-black nomargin" id="login-form-submit" name="login-form-submit" value="login">Login</button>--%>
                                                                </div>
                                                                <div class="col_full nobottommargin right">
                                                                    <asp:LinkButton ID="LinkButton_ForgotPassword" runat="server">Password Reset</asp:LinkButton>
                                                                    <asp:Literal ID="FailureText" runat="server" EnableViewState="False"></asp:Literal>
                                                                </div>
                                                            </form>

                                                        </LayoutTemplate>
                                                    </asp:Login>
                                                </AnonymousTemplate>
                                                <LoggedInTemplate>
                                                    <asp:LoginName ID="LoginName1" runat="server" />
                                                    <br />
                                                    <br />
                                                    <asp:Button ID="Button_logout" runat="server" Text="Log out" OnClick="Button_logout_Click" />

                                                </LoggedInTemplate>

                                            </asp:LoginView>

                                        </ContentTemplate>
                                    </asp:UpdatePanel>



                              
                            </div>

                            <div class="center dark"><small>Copyrights &copy; All Rights Reserved by Canvas Inc.</small></div>

                        </div>
                    </div>

                </div>

            </section>
            <!-- #content end -->

        </div>
        <!-- #wrapper end -->
    </form>

    <!-- Go To Top
	============================================= -->
    <div id="gotoTop" class="icon-angle-up"></div>

    <!-- External JavaScripts
	============================================= -->
    <script src="js/jquery.js"></script>
    <script src="js/plugins.js"></script>

    <!-- Footer Scripts
	============================================= -->
    <script src="js/functions.js"></script>

</body>
</html>
