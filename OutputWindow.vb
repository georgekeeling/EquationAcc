Public Class OutputWindow
    Private Sub BtnCopy_Click(sender As Object, e As EventArgs) Handles BtnCopyAll.Click
        My.Computer.Clipboard.SetText(ResultsBox.Text)
    End Sub

    Private Sub OutputWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OutputWindow_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim RBsize As New Drawing.Size With {.Height = Size.Height - 180, .Width = Size.Width - 60}  'suggested "impreovement"
        Dim Blocation As New Drawing.Point

        Blocation.X = Size.Width - 230
        Blocation.Y = Size.Height - 140
        BtnClose.Location = Blocation
        Blocation.X = Size.Width - 400
        BtnCopyAll.Location = Blocation
        ResultsBox.Size = RBsize
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Close()
    End Sub
End Class