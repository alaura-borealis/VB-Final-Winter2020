<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMultipleChoice
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMultipleChoice))
        Me.lblQuestion = New System.Windows.Forms.Label()
        Me.btnSubmit = New System.Windows.Forms.Button()
        Me.grpRadio = New System.Windows.Forms.Panel()
        Me.rad4 = New System.Windows.Forms.RadioButton()
        Me.rad3 = New System.Windows.Forms.RadioButton()
        Me.rad2 = New System.Windows.Forms.RadioButton()
        Me.rad1 = New System.Windows.Forms.RadioButton()
        Me.grpRadio.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblQuestion
        '
        Me.lblQuestion.AutoSize = True
        Me.lblQuestion.Location = New System.Drawing.Point(10, 20)
        Me.lblQuestion.Margin = New System.Windows.Forms.Padding(10, 10, 10, 15)
        Me.lblQuestion.MaximumSize = New System.Drawing.Size(290, 0)
        Me.lblQuestion.MinimumSize = New System.Drawing.Size(0, 30)
        Me.lblQuestion.Name = "lblQuestion"
        Me.lblQuestion.Padding = New System.Windows.Forms.Padding(3)
        Me.lblQuestion.Size = New System.Drawing.Size(6, 30)
        Me.lblQuestion.TabIndex = 0
        Me.lblQuestion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSubmit
        '
        Me.btnSubmit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSubmit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnSubmit.Location = New System.Drawing.Point(119, 177)
        Me.btnSubmit.Margin = New System.Windows.Forms.Padding(0, 15, 0, 15)
        Me.btnSubmit.Name = "btnSubmit"
        Me.btnSubmit.Size = New System.Drawing.Size(75, 23)
        Me.btnSubmit.TabIndex = 5
        Me.btnSubmit.Text = "Submit"
        Me.btnSubmit.UseVisualStyleBackColor = True
        '
        'grpRadio
        '
        Me.grpRadio.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.grpRadio.AutoSize = True
        Me.grpRadio.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.grpRadio.Controls.Add(Me.rad4)
        Me.grpRadio.Controls.Add(Me.rad3)
        Me.grpRadio.Controls.Add(Me.rad2)
        Me.grpRadio.Controls.Add(Me.rad1)
        Me.grpRadio.Location = New System.Drawing.Point(16, 55)
        Me.grpRadio.Margin = New System.Windows.Forms.Padding(10, 15, 10, 10)
        Me.grpRadio.Name = "grpRadio"
        Me.grpRadio.Size = New System.Drawing.Size(25, 103)
        Me.grpRadio.TabIndex = 6
        '
        'rad4
        '
        Me.rad4.AutoSize = True
        Me.rad4.Location = New System.Drawing.Point(8, 87)
        Me.rad4.Name = "rad4"
        Me.rad4.Size = New System.Drawing.Size(14, 13)
        Me.rad4.TabIndex = 4
        Me.rad4.TabStop = True
        Me.rad4.UseVisualStyleBackColor = True
        '
        'rad3
        '
        Me.rad3.AutoSize = True
        Me.rad3.Location = New System.Drawing.Point(8, 59)
        Me.rad3.Name = "rad3"
        Me.rad3.Size = New System.Drawing.Size(14, 13)
        Me.rad3.TabIndex = 3
        Me.rad3.TabStop = True
        Me.rad3.UseVisualStyleBackColor = True
        '
        'rad2
        '
        Me.rad2.AutoSize = True
        Me.rad2.Location = New System.Drawing.Point(8, 31)
        Me.rad2.Name = "rad2"
        Me.rad2.Size = New System.Drawing.Size(14, 13)
        Me.rad2.TabIndex = 2
        Me.rad2.TabStop = True
        Me.rad2.UseVisualStyleBackColor = True
        '
        'rad1
        '
        Me.rad1.AutoSize = True
        Me.rad1.Location = New System.Drawing.Point(8, 3)
        Me.rad1.Name = "rad1"
        Me.rad1.Size = New System.Drawing.Size(14, 13)
        Me.rad1.TabIndex = 1
        Me.rad1.TabStop = True
        Me.rad1.UseVisualStyleBackColor = True
        '
        'frmMultipleChoice
        '
        Me.AcceptButton = Me.btnSubmit
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(312, 225)
        Me.Controls.Add(Me.grpRadio)
        Me.Controls.Add(Me.btnSubmit)
        Me.Controls.Add(Me.lblQuestion)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMultipleChoice"
        Me.Padding = New System.Windows.Forms.Padding(0, 0, 0, 25)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Quiz"
        Me.grpRadio.ResumeLayout(False)
        Me.grpRadio.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblQuestion As Label
    Friend WithEvents btnSubmit As Button
    Friend WithEvents grpRadio As Panel
    Friend WithEvents rad4 As RadioButton
    Friend WithEvents rad3 As RadioButton
    Friend WithEvents rad2 As RadioButton
    Friend WithEvents rad1 As RadioButton
End Class
