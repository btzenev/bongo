<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.txt_pass = New System.Windows.Forms.TextBox()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.dlg_file = New System.Windows.Forms.OpenFileDialog()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.lbl_file = New System.Windows.Forms.Label()
    Me.txt_id = New System.Windows.Forms.TextBox()
    Me.btn_clear = New System.Windows.Forms.Button()
    Me.btn_gen = New System.Windows.Forms.Button()
    Me.lst_keys = New System.Windows.Forms.ListBox()
    Me.txt_key = New System.Windows.Forms.TextBox()
    Me.btn_add = New System.Windows.Forms.Button()
    Me.SuspendLayout()
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(5, 64)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(54, 13)
    Me.Label3.TabIndex = 16
    Me.Label3.Text = "Site Keys:"
    '
    'txt_pass
    '
    Me.txt_pass.Location = New System.Drawing.Point(62, 34)
    Me.txt_pass.Name = "txt_pass"
    Me.txt_pass.Size = New System.Drawing.Size(327, 20)
    Me.txt_pass.TabIndex = 15
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(5, 11)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(42, 13)
    Me.Label2.TabIndex = 11
    Me.Label2.Text = "Site ID:"
    '
    'dlg_file
    '
    Me.dlg_file.FileName = "web.config"
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(5, 37)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(57, 13)
    Me.Label1.TabIndex = 14
    Me.Label1.Text = "Site Pass: "
    '
    'lbl_file
    '
    Me.lbl_file.AutoSize = True
    Me.lbl_file.BackColor = System.Drawing.Color.White
    Me.lbl_file.Location = New System.Drawing.Point(369, 11)
    Me.lbl_file.Name = "lbl_file"
    Me.lbl_file.Size = New System.Drawing.Size(16, 13)
    Me.lbl_file.TabIndex = 13
    Me.lbl_file.Text = "..."
    '
    'txt_id
    '
    Me.txt_id.Location = New System.Drawing.Point(62, 7)
    Me.txt_id.Name = "txt_id"
    Me.txt_id.Size = New System.Drawing.Size(327, 20)
    Me.txt_id.TabIndex = 12
    '
    'btn_clear
    '
    Me.btn_clear.Location = New System.Drawing.Point(6, 202)
    Me.btn_clear.Name = "btn_clear"
    Me.btn_clear.Size = New System.Drawing.Size(83, 23)
    Me.btn_clear.TabIndex = 20
    Me.btn_clear.Text = "Clear All Keys"
    Me.btn_clear.UseVisualStyleBackColor = True
    '
    'btn_gen
    '
    Me.btn_gen.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.btn_gen.Location = New System.Drawing.Point(209, 202)
    Me.btn_gen.Name = "btn_gen"
    Me.btn_gen.Size = New System.Drawing.Size(180, 37)
    Me.btn_gen.TabIndex = 21
    Me.btn_gen.Text = "&Generate"
    Me.btn_gen.UseVisualStyleBackColor = True
    '
    'lst_keys
    '
    Me.lst_keys.FormattingEnabled = True
    Me.lst_keys.Location = New System.Drawing.Point(6, 88)
    Me.lst_keys.Name = "lst_keys"
    Me.lst_keys.Size = New System.Drawing.Size(383, 108)
    Me.lst_keys.TabIndex = 19
    '
    'txt_key
    '
    Me.txt_key.Location = New System.Drawing.Point(62, 61)
    Me.txt_key.Name = "txt_key"
    Me.txt_key.Size = New System.Drawing.Size(267, 20)
    Me.txt_key.TabIndex = 17
    '
    'btn_add
    '
    Me.btn_add.Location = New System.Drawing.Point(335, 59)
    Me.btn_add.Name = "btn_add"
    Me.btn_add.Size = New System.Drawing.Size(54, 23)
    Me.btn_add.TabIndex = 18
    Me.btn_add.Text = "Add"
    Me.btn_add.UseVisualStyleBackColor = True
    '
    'Main
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(395, 246)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.txt_pass)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.lbl_file)
    Me.Controls.Add(Me.txt_id)
    Me.Controls.Add(Me.btn_clear)
    Me.Controls.Add(Me.btn_gen)
    Me.Controls.Add(Me.lst_keys)
    Me.Controls.Add(Me.txt_key)
    Me.Controls.Add(Me.btn_add)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.Name = "Main"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "License Generator"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents txt_pass As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents dlg_file As System.Windows.Forms.OpenFileDialog
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents lbl_file As System.Windows.Forms.Label
  Friend WithEvents txt_id As System.Windows.Forms.TextBox
  Friend WithEvents btn_clear As System.Windows.Forms.Button
  Friend WithEvents btn_gen As System.Windows.Forms.Button
  Friend WithEvents lst_keys As System.Windows.Forms.ListBox
  Friend WithEvents txt_key As System.Windows.Forms.TextBox
  Friend WithEvents btn_add As System.Windows.Forms.Button

End Class
