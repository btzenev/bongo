Imports System.Drawing
Imports System.Configuration
Imports System.Web.Configuration

Public Class Main

  Private mstr_config As String
  Private mstr_lic As String

  Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
    If Not lst_keys.Items.Contains(txt_key.Text) Then lst_keys.Items.Add(txt_key.Text)
    txt_key.Text = ""
  End Sub

  Private Sub btn_clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_clear.Click
    If MessageBox.Show("Are you sure?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
      lst_keys.Items.Clear()
      txt_key.Text = ""
      txt_key.Focus()
    End If
  End Sub

  Private Sub btn_gen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_gen.Click
    If lst_keys.Items.Count = 0 Then
      MessageBox.Show("Please add at least one license")
      Return
    End If

    If IO.File.Exists(mstr_lic) Then
      If Not MessageBox.Show("Do you want to override the existing license?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then Return
      IO.File.Delete(mstr_lic)
    End If
    Dim obj_list As New List(Of String)
    obj_list.Add("pass:" & txt_pass.Text.Replace(" ", "").Trim())
    For Each str_value As String In lst_keys.Items
      If Not String.IsNullOrEmpty(str_value) Then obj_list.Add(str_value)
    Next
    Dim str_data As String = Join(obj_list.ToArray, vbNewLine)

    Dim obj_key As New Encryption.Data(txt_id.Text)
    Dim obj_data As New Encryption.Data(str_data)
    Dim obj_syn As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)

    obj_data = obj_syn.Encrypt(obj_data, obj_key)

    Using obj_file As New IO.FileStream(mstr_lic, IO.FileMode.Create)
      Using obj_writer As New IO.StreamWriter(obj_file)
        obj_writer.Write(obj_data.ToHex)
      End Using
    End Using

    'resave config file to update force latest changes
    Dim obj_map As New ExeConfigurationFileMap()
    obj_map.ExeConfigFilename = mstr_config
    Dim obj_config As Configuration = ConfigurationManager.OpenMappedExeConfiguration(obj_map, ConfigurationUserLevel.None)
    obj_config.Save()
  End Sub

  Private Sub txt_key_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_key.KeyDown
    If e.KeyCode = Keys.Enter Then btn_add_Click(sender, e)
  End Sub

  Private Sub LoadConfig()
    'On Error Resume Next
    dlg_file.Filter = "Config|web.config"
    dlg_file.AddExtension = True
    dlg_file.InitialDirectory = IO.Path.GetDirectoryName(Windows.Forms.Application.ExecutablePath)
    If dlg_file.ShowDialog() = DialogResult.OK Then

      Dim obj_map As New ExeConfigurationFileMap()
      mstr_config = dlg_file.FileName
      obj_map.ExeConfigFilename = mstr_config

      Dim obj_config As Configuration = ConfigurationManager.OpenMappedExeConfiguration(obj_map, ConfigurationUserLevel.None)
      Dim obj_settings As AppSettingsSection = DirectCast(obj_config.GetSection("appSettings"), AppSettingsSection)
      If obj_settings IsNot Nothing Then
        Dim str_id As String = String.Empty
        If obj_settings.Settings("cms.id") IsNot Nothing Then
          str_id = obj_settings.Settings("cms.id").Value.Trim()
        End If
        If String.IsNullOrEmpty(str_id) Then
          If MsgBox("No ID found. Generate one?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Try
              str_id = Guid.NewGuid.ToString.TrimStart("{").TrimEnd("}").ToUpper
              obj_settings.Settings.Remove("cms.id")
              obj_settings.Settings.Add("cms.id", str_id)
              obj_config.Save()
            Catch ex As Exception
              MsgBox(ex.Message)
              End
            End Try
          Else
            End
          End If
        End If
        txt_id.Text = str_id

        mstr_lic = IO.Path.Combine(IO.Path.GetDirectoryName(mstr_config), "license.dat")
        If IO.File.Exists(mstr_lic) Then
          Dim str_lic As String = ReadFile(mstr_lic).Trim
          Dim obj_key As New Encryption.Data(str_id)
          Dim obj_data As New Encryption.Data()
          obj_data.Hex = str_lic
          Dim obj_syn As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
          obj_data = obj_syn.Decrypt(obj_data, obj_key)

          'Dim obj_list As String() = 
          Dim bln_first As Boolean = True
          For Each str_key As String In obj_data.ToString.ToLower.Split(vbNewLine)
            If bln_first And str_key.Contains("pass:") Then
              txt_pass.Text = str_key.Replace("pass:", "")
            Else
              lst_keys.Items.Add(str_key)
            End If
            bln_first = False
          Next
          'lst_keys.Items.AddRange(obj_data.ToString.ToLower.Split(vbNewLine))
        End If
      Else
        MsgBox("No appSettings found.")
        End
      End If

      txt_key.Focus()
    Else
      End
    End If
  End Sub

  Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    On Error Resume Next
    LoadConfig()
  End Sub

  Private Sub lbl_file_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbl_file.Click
    On Error Resume Next
    LoadConfig()
  End Sub

  Private Sub lbl_file_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbl_file.MouseHover
    lbl_file.BorderStyle = BorderStyle.Fixed3D
  End Sub

  Private Sub lbl_file_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbl_file.MouseLeave
    lbl_file.BorderStyle = BorderStyle.None
  End Sub

  Private Sub txt_id_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt_id.DoubleClick
    txt_id.Text = Guid.NewGuid.ToString.ToUpper
    txt_key.Focus()
  End Sub

  Private Sub dlg_file_FileOk(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles dlg_file.FileOk

  End Sub

End Class
