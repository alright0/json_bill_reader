Imports System.Security.Cryptography
Imports System.Text

Public Class Form3
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        OpenFileDialog1.Filter = "license files (*.lic)|*.lic"

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim lic As String = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            My.Computer.FileSystem.WriteAllText(Application.StartupPath & "\license.lic", lic, False)

            ' проверка состояния после загрузки лицензии
            If Form1.lic_compare() Then
                MsgBox("Файл успешно загружен!" & Chr(10) & "Перезапустите приложение")
            Else
                MsgBox("Загружен недействительный файл лицензии!")

            End If


        End If

    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim FSet = My.Computer.Registry.GetValue(
            "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "BaseBoardProduct", Nothing)

        Dim Sset = My.Computer.Registry.GetValue(
            "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString", Nothing)

        Dim Dset = FSet & " " & Sset


        TextBox5.Text = Form1.GetHash(Dset)
    End Sub
End Class