Public Class _login
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button_logout_Click(sender As Object, e As EventArgs)
        FormsAuthentication.SignOut()
        'UpdatePanel1.Update()
        Response.Redirect(HttpContext.Current.Request.Url.ToString(), True)
    End Sub
End Class