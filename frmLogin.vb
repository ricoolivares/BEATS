Public Class frmLogin

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'close Login form.
        Me.Close()
    End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim errorMsg As String
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) \ 2

        'Getting actual version       
        lblVersion.Text = "Version " + My.Application.Info.Version.ToString


        'Opening configuration file and initialization an creating SETTING's class object (miConfiguracion)
        errorMsg = openConfigurationFile()

        'Checking if an error ocurrs

        If errorMsg <> "" Then

            'Display an error message
            MsgBox(errorMsg, MsgBoxStyle.Critical)

            'End the application
            End

        End If

        txtPassword.Focus()

    End Sub

    Private Sub txtPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyDown
        Dim user, pass, errorMsg As String

        If e.KeyCode = Keys.Enter Then

            user = Trim(txtUser.Text)
            pass = Trim(txtPassword.Text)

            If UCase(user) = "ADMIN" And pass = "Dr3amer" Then
                errorMsg = ""
            Else
                errorMsg = "!Credenciales Invalidas!"
            End If



            'If pass credential
            If errorMsg = "" Then


                'Hide login form
                Me.Hide()

                
                'Show Maintenance form
                'frmLoadBeats.Show()

            Else
                'Displaying error
                lblMessage.Visible = True
                lblMessage.Text = errorMsg
            End If

            txtPassword.Text = ""

        End If

    End Sub
     
    
End Class
