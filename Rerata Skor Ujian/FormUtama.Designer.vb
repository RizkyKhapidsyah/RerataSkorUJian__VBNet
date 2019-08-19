<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormUtama
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.labelRerata = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.teksSkor3 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.teksSkor2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.teksSkor1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.labelPesan = New System.Windows.Forms.Label()
        Me.tombolHitung = New System.Windows.Forms.Button()
        Me.tombolBersihkan = New System.Windows.Forms.Button()
        Me.tombolKeluar = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.labelRerata)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.teksSkor3)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.teksSkor2)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.teksSkor1)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(43, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(193, 176)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Masukkan Tiga Skor Ujian"
        '
        'labelRerata
        '
        Me.labelRerata.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.labelRerata.Location = New System.Drawing.Point(59, 139)
        Me.labelRerata.Name = "labelRerata"
        Me.labelRerata.Size = New System.Drawing.Size(113, 23)
        Me.labelRerata.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 142)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(42, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Rerata:"
        '
        'teksSkor3
        '
        Me.teksSkor3.Location = New System.Drawing.Point(59, 91)
        Me.teksSkor3.Name = "teksSkor3"
        Me.teksSkor3.Size = New System.Drawing.Size(113, 20)
        Me.teksSkor3.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 94)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Skor 3:"
        '
        'teksSkor2
        '
        Me.teksSkor2.Location = New System.Drawing.Point(59, 50)
        Me.teksSkor2.Name = "teksSkor2"
        Me.teksSkor2.Size = New System.Drawing.Size(113, 20)
        Me.teksSkor2.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Skor 2:"
        '
        'teksSkor1
        '
        Me.teksSkor1.Location = New System.Drawing.Point(59, 17)
        Me.teksSkor1.Name = "teksSkor1"
        Me.teksSkor1.Size = New System.Drawing.Size(113, 20)
        Me.teksSkor1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Skor 1:"
        '
        'labelPesan
        '
        Me.labelPesan.Location = New System.Drawing.Point(43, 202)
        Me.labelPesan.Name = "labelPesan"
        Me.labelPesan.Size = New System.Drawing.Size(193, 23)
        Me.labelPesan.TabIndex = 1
        Me.labelPesan.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tombolHitung
        '
        Me.tombolHitung.Location = New System.Drawing.Point(43, 244)
        Me.tombolHitung.Name = "tombolHitung"
        Me.tombolHitung.Size = New System.Drawing.Size(96, 60)
        Me.tombolHitung.TabIndex = 2
        Me.tombolHitung.Text = "Hitung Rerata"
        Me.tombolHitung.UseVisualStyleBackColor = True
        '
        'tombolBersihkan
        '
        Me.tombolBersihkan.Location = New System.Drawing.Point(145, 244)
        Me.tombolBersihkan.Name = "tombolBersihkan"
        Me.tombolBersihkan.Size = New System.Drawing.Size(91, 29)
        Me.tombolBersihkan.TabIndex = 3
        Me.tombolBersihkan.Text = "Bersihkan"
        Me.tombolBersihkan.UseVisualStyleBackColor = True
        '
        'tombolKeluar
        '
        Me.tombolKeluar.Location = New System.Drawing.Point(146, 276)
        Me.tombolKeluar.Name = "tombolKeluar"
        Me.tombolKeluar.Size = New System.Drawing.Size(90, 29)
        Me.tombolKeluar.TabIndex = 4
        Me.tombolKeluar.Text = "Keluar"
        Me.tombolKeluar.UseVisualStyleBackColor = True
        '
        'FormUtama
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(279, 313)
        Me.Controls.Add(Me.tombolKeluar)
        Me.Controls.Add(Me.tombolBersihkan)
        Me.Controls.Add(Me.tombolHitung)
        Me.Controls.Add(Me.labelPesan)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "FormUtama"
        Me.Text = "Rerata Skor Ujian"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents labelRerata As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents teksSkor3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents teksSkor2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents teksSkor1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents labelPesan As System.Windows.Forms.Label
    Friend WithEvents tombolHitung As System.Windows.Forms.Button
    Friend WithEvents tombolBersihkan As System.Windows.Forms.Button
    Friend WithEvents tombolKeluar As System.Windows.Forms.Button

End Class
