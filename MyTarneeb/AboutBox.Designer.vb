<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AboutBox
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    Friend WithEvents TableLayoutPanel As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents LogoPictureBox As System.Windows.Forms.PictureBox
    Friend WithEvents LabelProductName As System.Windows.Forms.Label
    Friend WithEvents LabelVersion As System.Windows.Forms.Label
    Friend WithEvents LabelCompanyName As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents LabelCopyright As System.Windows.Forms.Label

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AboutBox))
        Me.TableLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
        Me.LabelProductName = New System.Windows.Forms.Label()
        Me.LabelVersion = New System.Windows.Forms.Label()
        Me.LabelCopyright = New System.Windows.Forms.Label()
        Me.LabelCompanyName = New System.Windows.Forms.Label()
        Me.LogoPictureBox = New System.Windows.Forms.PictureBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.SpecialThanksLabel = New System.Windows.Forms.Label()
        Me.DevLabel = New System.Windows.Forms.Label()
        Me.CardsDesignLinkLabel = New System.Windows.Forms.LinkLabel()
        Me.TableLayoutPanel.SuspendLayout()
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel
        '
        Me.TableLayoutPanel.ColumnCount = 2
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120.0!))
        Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel.Controls.Add(Me.LabelProductName, 1, 0)
        Me.TableLayoutPanel.Controls.Add(Me.LabelVersion, 1, 1)
        Me.TableLayoutPanel.Controls.Add(Me.LabelCopyright, 1, 2)
        Me.TableLayoutPanel.Controls.Add(Me.LabelCompanyName, 1, 3)
        Me.TableLayoutPanel.Controls.Add(Me.LogoPictureBox, 0, 0)
        Me.TableLayoutPanel.Controls.Add(Me.OKButton, 1, 6)
        Me.TableLayoutPanel.Controls.Add(Me.SpecialThanksLabel, 1, 5)
        Me.TableLayoutPanel.Controls.Add(Me.DevLabel, 1, 4)
        Me.TableLayoutPanel.Controls.Add(Me.CardsDesignLinkLabel, 0, 6)
        Me.TableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel.Location = New System.Drawing.Point(9, 9)
        Me.TableLayoutPanel.Name = "TableLayoutPanel"
        Me.TableLayoutPanel.RowCount = 7
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.39437!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.39437!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.39437!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.39437!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.39597!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 37.63488!))
        Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.39167!))
        Me.TableLayoutPanel.Size = New System.Drawing.Size(564, 258)
        Me.TableLayoutPanel.TabIndex = 0
        '
        'LabelProductName
        '
        Me.LabelProductName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelProductName.Location = New System.Drawing.Point(126, 0)
        Me.LabelProductName.Margin = New System.Windows.Forms.Padding(6, 0, 3, 0)
        Me.LabelProductName.MaximumSize = New System.Drawing.Size(0, 17)
        Me.LabelProductName.Name = "LabelProductName"
        Me.LabelProductName.Size = New System.Drawing.Size(435, 17)
        Me.LabelProductName.TabIndex = 0
        Me.LabelProductName.Text = "Nom du produit"
        Me.LabelProductName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelVersion
        '
        Me.LabelVersion.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelVersion.Location = New System.Drawing.Point(126, 26)
        Me.LabelVersion.Margin = New System.Windows.Forms.Padding(6, 0, 3, 0)
        Me.LabelVersion.MaximumSize = New System.Drawing.Size(0, 17)
        Me.LabelVersion.Name = "LabelVersion"
        Me.LabelVersion.Size = New System.Drawing.Size(435, 17)
        Me.LabelVersion.TabIndex = 0
        Me.LabelVersion.Text = "Version"
        Me.LabelVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelCopyright
        '
        Me.LabelCopyright.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelCopyright.Location = New System.Drawing.Point(126, 52)
        Me.LabelCopyright.Margin = New System.Windows.Forms.Padding(6, 0, 3, 0)
        Me.LabelCopyright.MaximumSize = New System.Drawing.Size(0, 17)
        Me.LabelCopyright.Name = "LabelCopyright"
        Me.LabelCopyright.Size = New System.Drawing.Size(435, 17)
        Me.LabelCopyright.TabIndex = 0
        Me.LabelCopyright.Text = "Copyright"
        Me.LabelCopyright.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelCompanyName
        '
        Me.LabelCompanyName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelCompanyName.Location = New System.Drawing.Point(126, 78)
        Me.LabelCompanyName.Margin = New System.Windows.Forms.Padding(6, 0, 3, 0)
        Me.LabelCompanyName.MaximumSize = New System.Drawing.Size(0, 17)
        Me.LabelCompanyName.Name = "LabelCompanyName"
        Me.LabelCompanyName.Size = New System.Drawing.Size(435, 17)
        Me.LabelCompanyName.TabIndex = 0
        Me.LabelCompanyName.Text = "Nom de la société"
        Me.LabelCompanyName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LogoPictureBox
        '
        Me.LogoPictureBox.Image = CType(resources.GetObject("LogoPictureBox.Image"), System.Drawing.Image)
        Me.LogoPictureBox.Location = New System.Drawing.Point(3, 3)
        Me.LogoPictureBox.Name = "LogoPictureBox"
        Me.TableLayoutPanel.SetRowSpan(Me.LogoPictureBox, 5)
        Me.LogoPictureBox.Size = New System.Drawing.Size(114, 124)
        Me.LogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.LogoPictureBox.TabIndex = 0
        Me.LogoPictureBox.TabStop = False
        '
        'OKButton
        '
        Me.OKButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.OKButton.Location = New System.Drawing.Point(486, 232)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(75, 23)
        Me.OKButton.TabIndex = 0
        Me.OKButton.Text = "&OK"
        '
        'SpecialThanksLabel
        '
        Me.SpecialThanksLabel.AutoSize = True
        Me.SpecialThanksLabel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SpecialThanksLabel.Location = New System.Drawing.Point(123, 130)
        Me.SpecialThanksLabel.Name = "SpecialThanksLabel"
        Me.SpecialThanksLabel.Size = New System.Drawing.Size(438, 97)
        Me.SpecialThanksLabel.TabIndex = 1
        Me.SpecialThanksLabel.Text = "Remerciements à :"
        Me.SpecialThanksLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DevLabel
        '
        Me.DevLabel.AutoSize = True
        Me.DevLabel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DevLabel.Location = New System.Drawing.Point(123, 104)
        Me.DevLabel.MaximumSize = New System.Drawing.Size(0, 17)
        Me.DevLabel.Name = "DevLabel"
        Me.DevLabel.Size = New System.Drawing.Size(438, 17)
        Me.DevLabel.TabIndex = 2
        Me.DevLabel.Text = "Développé par"
        Me.DevLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CardsDesignLinkLabel
        '
        Me.CardsDesignLinkLabel.AutoSize = True
        Me.CardsDesignLinkLabel.Location = New System.Drawing.Point(3, 227)
        Me.CardsDesignLinkLabel.Name = "CardsDesignLinkLabel"
        Me.CardsDesignLinkLabel.Size = New System.Drawing.Size(100, 26)
        Me.CardsDesignLinkLabel.TabIndex = 3
        Me.CardsDesignLinkLabel.TabStop = True
        Me.CardsDesignLinkLabel.Text = "Site web du design des cartes"
        '
        'AboutBox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.OKButton
        Me.ClientSize = New System.Drawing.Size(582, 276)
        Me.Controls.Add(Me.TableLayoutPanel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AboutBox"
        Me.Padding = New System.Windows.Forms.Padding(9)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "AboutBox"
        Me.TableLayoutPanel.ResumeLayout(False)
        Me.TableLayoutPanel.PerformLayout()
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SpecialThanksLabel As System.Windows.Forms.Label
    Friend WithEvents DevLabel As System.Windows.Forms.Label
    Friend WithEvents CardsDesignLinkLabel As System.Windows.Forms.LinkLabel

End Class
