<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmGameBoard
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGameBoard))
        Me.btnP1 = New System.Windows.Forms.Button()
        Me.btnP2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnP1
        '
        Me.btnP1.Location = New System.Drawing.Point(282, 629)
        Me.btnP1.Name = "btnP1"
        Me.btnP1.Size = New System.Drawing.Size(75, 23)
        Me.btnP1.TabIndex = 0
        Me.btnP1.Text = "Player 1"
        Me.btnP1.UseVisualStyleBackColor = True
        '
        'btnP2
        '
        Me.btnP2.Location = New System.Drawing.Point(442, 629)
        Me.btnP2.Name = "btnP2"
        Me.btnP2.Size = New System.Drawing.Size(75, 23)
        Me.btnP2.TabIndex = 1
        Me.btnP2.Text = "Player 2"
        Me.btnP2.UseVisualStyleBackColor = True
        '
        'frmGameBoard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.EnvBoardGame.My.Resources.Resources.game_board
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(800, 801)
        Me.Controls.Add(Me.btnP2)
        Me.Controls.Add(Me.btnP1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmGameBoard"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Roots & Branches"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnP1 As Button
    Friend WithEvents btnP2 As Button
End Class
