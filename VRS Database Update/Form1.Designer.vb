<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        TextBox1 = New TextBox()
        Button1 = New Button()
        TextBox2 = New TextBox()
        TextBox3 = New TextBox()
        TextBox4 = New TextBox()
        TextBox5 = New TextBox()
        TextBox6 = New TextBox()
        TextBox7 = New TextBox()
        TextBox8 = New TextBox()
        TextBox9 = New TextBox()
        TextBox10 = New TextBox()
        CheckBox1 = New CheckBox()
        CheckBox2 = New CheckBox()
        Button2 = New Button()
        Button3 = New Button()
        SettingsToolStripMenuItem = New ToolStripMenuItem()
        ExitToolStripMenuItem = New ToolStripMenuItem()
        HelpToolStripMenuItem = New ToolStripMenuItem()
        MenuStrip1 = New MenuStrip()
        CheckBox3 = New CheckBox()
        TextBox11 = New TextBox()
        Button4 = New Button()
        CheckBox5 = New CheckBox()
        CheckBox4 = New CheckBox()
        MenuStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New Point(12, 96)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(336, 193)
        TextBox1.TabIndex = 0
        ' 
        ' Button1
        ' 
        Button1.BackColor = Color.Lime
        Button1.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        Button1.Location = New Point(355, 96)
        Button1.Name = "Button1"
        Button1.Size = New Size(131, 33)
        Button1.TabIndex = 1
        Button1.Text = "Start"
        Button1.UseVisualStyleBackColor = False
        ' 
        ' TextBox2
        ' 
        TextBox2.BackColor = SystemColors.WindowText
        TextBox2.Font = New Font("Segoe UI", 20F, FontStyle.Bold, GraphicsUnit.Point)
        TextBox2.ForeColor = Color.Red
        TextBox2.Location = New Point(12, 26)
        TextBox2.Name = "TextBox2"
        TextBox2.Size = New Size(511, 43)
        TextBox2.TabIndex = 2
        TextBox2.Text = "E's VRS Database Updater"
        TextBox2.TextAlign = HorizontalAlignment.Center
        ' 
        ' TextBox3
        ' 
        TextBox3.BackColor = SystemColors.Menu
        TextBox3.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        TextBox3.Location = New Point(122, 72)
        TextBox3.Name = "TextBox3"
        TextBox3.Size = New Size(113, 23)
        TextBox3.TabIndex = 3
        TextBox3.Text = "Program Progress"
        TextBox3.TextAlign = HorizontalAlignment.Center
        ' 
        ' TextBox4
        ' 
        TextBox4.Location = New Point(12, 319)
        TextBox4.Multiline = True
        TextBox4.Name = "TextBox4"
        TextBox4.Size = New Size(511, 301)
        TextBox4.TabIndex = 4
        ' 
        ' TextBox5
        ' 
        TextBox5.BackColor = SystemColors.Menu
        TextBox5.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        TextBox5.Location = New Point(189, 295)
        TextBox5.Name = "TextBox5"
        TextBox5.Size = New Size(159, 23)
        TextBox5.TabIndex = 5
        TextBox5.Text = "Database Update Status"
        TextBox5.TextAlign = HorizontalAlignment.Center
        ' 
        ' TextBox6
        ' 
        TextBox6.BackColor = SystemColors.Menu
        TextBox6.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        TextBox6.Location = New Point(355, 242)
        TextBox6.Name = "TextBox6"
        TextBox6.Size = New Size(51, 23)
        TextBox6.TabIndex = 6
        TextBox6.Text = " - %"
        TextBox6.TextAlign = HorizontalAlignment.Right
        ' 
        ' TextBox7
        ' 
        TextBox7.BackColor = SystemColors.Menu
        TextBox7.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        TextBox7.Location = New Point(355, 265)
        TextBox7.Name = "TextBox7"
        TextBox7.Size = New Size(51, 23)
        TextBox7.TabIndex = 7
        TextBox7.Text = " - "
        TextBox7.TextAlign = HorizontalAlignment.Right
        ' 
        ' TextBox8
        ' 
        TextBox8.BackColor = SystemColors.Menu
        TextBox8.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        TextBox8.Location = New Point(405, 265)
        TextBox8.Name = "TextBox8"
        TextBox8.Size = New Size(119, 23)
        TextBox8.TabIndex = 8
        TextBox8.Text = "Time Remaining"
        ' 
        ' TextBox9
        ' 
        TextBox9.BackColor = SystemColors.Menu
        TextBox9.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        TextBox9.Location = New Point(406, 242)
        TextBox9.Name = "TextBox9"
        TextBox9.Size = New Size(80, 23)
        TextBox9.TabIndex = 9
        TextBox9.Text = "Completed"
        ' 
        ' TextBox10
        ' 
        TextBox10.Location = New Point(12, 626)
        TextBox10.Multiline = True
        TextBox10.Name = "TextBox10"
        TextBox10.Size = New Size(407, 44)
        TextBox10.TabIndex = 10
        ' 
        ' CheckBox1
        ' 
        CheckBox1.AutoSize = True
        CheckBox1.Location = New Point(12, 726)
        CheckBox1.Name = "CheckBox1"
        CheckBox1.Size = New Size(116, 19)
        CheckBox1.TabIndex = 13
        CheckBox1.Text = "Run immediately"
        CheckBox1.UseVisualStyleBackColor = True
        ' 
        ' CheckBox2
        ' 
        CheckBox2.AutoSize = True
        CheckBox2.Location = New Point(12, 751)
        CheckBox2.Name = "CheckBox2"
        CheckBox2.Size = New Size(208, 19)
        CheckBox2.TabIndex = 13
        CheckBox2.Text = "Always re-download FAA database"
        CheckBox2.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.BackColor = SystemColors.ActiveCaption
        Button2.Location = New Point(425, 626)
        Button2.Name = "Button2"
        Button2.Size = New Size(98, 44)
        Button2.TabIndex = 14
        Button2.Text = "Choose working folder"
        Button2.UseVisualStyleBackColor = False
        ' 
        ' Button3
        ' 
        Button3.BackColor = Color.MediumTurquoise
        Button3.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        Button3.Location = New Point(425, 726)
        Button3.Name = "Button3"
        Button3.Size = New Size(98, 40)
        Button3.TabIndex = 1
        Button3.Text = "CCAR database webpage"
        Button3.UseVisualStyleBackColor = False
        ' 
        ' SettingsToolStripMenuItem
        ' 
        SettingsToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ExitToolStripMenuItem})
        SettingsToolStripMenuItem.Name = "SettingsToolStripMenuItem"
        SettingsToolStripMenuItem.Size = New Size(37, 20)
        SettingsToolStripMenuItem.Text = "File"
        ' 
        ' ExitToolStripMenuItem
        ' 
        ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        ExitToolStripMenuItem.Size = New Size(93, 22)
        ExitToolStripMenuItem.Text = "Exit"
        ' 
        ' HelpToolStripMenuItem
        ' 
        HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        HelpToolStripMenuItem.Size = New Size(44, 20)
        HelpToolStripMenuItem.Text = "Help"
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {SettingsToolStripMenuItem, HelpToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(535, 24)
        MenuStrip1.TabIndex = 11
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' CheckBox3
        ' 
        CheckBox3.AutoSize = True
        CheckBox3.Location = New Point(268, 751)
        CheckBox3.Name = "CheckBox3"
        CheckBox3.Size = New Size(138, 19)
        CheckBox3.TabIndex = 15
        CheckBox3.Text = "Backup VRS database"
        CheckBox3.UseVisualStyleBackColor = True
        ' 
        ' TextBox11
        ' 
        TextBox11.Location = New Point(12, 676)
        TextBox11.Multiline = True
        TextBox11.Name = "TextBox11"
        TextBox11.Size = New Size(407, 44)
        TextBox11.TabIndex = 10
        ' 
        ' Button4
        ' 
        Button4.BackColor = SystemColors.ActiveCaption
        Button4.Location = New Point(425, 676)
        Button4.Name = "Button4"
        Button4.Size = New Size(98, 44)
        Button4.TabIndex = 14
        Button4.Text = "Choose VRS data folder"
        Button4.UseVisualStyleBackColor = False
        ' 
        ' CheckBox5
        ' 
        CheckBox5.AutoSize = True
        CheckBox5.Location = New Point(12, 776)
        CheckBox5.Name = "CheckBox5"
        CheckBox5.Size = New Size(234, 19)
        CheckBox5.TabIndex = 18
        CheckBox5.Text = "Always re-download OpenSky database"
        CheckBox5.UseVisualStyleBackColor = True
        ' 
        ' CheckBox4
        ' 
        CheckBox4.AutoSize = True
        CheckBox4.Location = New Point(268, 776)
        CheckBox4.Name = "CheckBox4"
        CheckBox4.Size = New Size(156, 19)
        CheckBox4.TabIndex = 19
        CheckBox4.Text = "Build complete database"
        CheckBox4.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(535, 795)
        Controls.Add(CheckBox4)
        Controls.Add(CheckBox5)
        Controls.Add(CheckBox3)
        Controls.Add(Button4)
        Controls.Add(Button2)
        Controls.Add(CheckBox2)
        Controls.Add(CheckBox1)
        Controls.Add(TextBox11)
        Controls.Add(TextBox10)
        Controls.Add(TextBox9)
        Controls.Add(TextBox8)
        Controls.Add(TextBox7)
        Controls.Add(TextBox6)
        Controls.Add(TextBox5)
        Controls.Add(TextBox4)
        Controls.Add(TextBox3)
        Controls.Add(TextBox2)
        Controls.Add(Button3)
        Controls.Add(Button1)
        Controls.Add(TextBox1)
        Controls.Add(MenuStrip1)
        MainMenuStrip = MenuStrip1
        Name = "Form1"
        Text = "E's VRS Updater v1.1"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()

    End Sub

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents TextBox6 As TextBox
    Friend WithEvents TextBox7 As TextBox
    Friend WithEvents TextBox8 As TextBox
    Friend WithEvents TextBox9 As TextBox
    Friend WithEvents TextBox10 As TextBox
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents SettingsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents TextBox11 As TextBox
    Friend WithEvents Button4 As Button
    Friend WithEvents CheckBox5 As CheckBox
    Friend WithEvents CheckBox4 As CheckBox
End Class
