<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form8
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.HOMEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GOTOToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SIGNUPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LOGOUTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ADMINToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.POSSYSTEMToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductInventoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CustomerListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DASHBOARDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Playbill", 72.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(323, 100)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(868, 98)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = " AGRI POINT OF SALE SYSTEM "
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.ForestGreen
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HOMEToolStripMenuItem, Me.ADMINToolStripMenuItem, Me.DASHBOARDToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(4, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(1370, 38)
        Me.MenuStrip1.TabIndex = 7
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'HOMEToolStripMenuItem
        '
        Me.HOMEToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GOTOToolStripMenuItem, Me.LOGOUTToolStripMenuItem})
        Me.HOMEToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HOMEToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.HOMEToolStripMenuItem.Name = "HOMEToolStripMenuItem"
        Me.HOMEToolStripMenuItem.Size = New System.Drawing.Size(89, 34)
        Me.HOMEToolStripMenuItem.Text = "HOME"
        '
        'GOTOToolStripMenuItem
        '
        Me.GOTOToolStripMenuItem.BackColor = System.Drawing.Color.LightYellow
        Me.GOTOToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SIGNUPToolStripMenuItem})
        Me.GOTOToolStripMenuItem.Name = "GOTOToolStripMenuItem"
        Me.GOTOToolStripMenuItem.Size = New System.Drawing.Size(177, 34)
        Me.GOTOToolStripMenuItem.Text = "GOTO"
        '
        'SIGNUPToolStripMenuItem
        '
        Me.SIGNUPToolStripMenuItem.Name = "SIGNUPToolStripMenuItem"
        Me.SIGNUPToolStripMenuItem.Size = New System.Drawing.Size(168, 34)
        Me.SIGNUPToolStripMenuItem.Text = "SIGN UP"
        '
        'LOGOUTToolStripMenuItem
        '
        Me.LOGOUTToolStripMenuItem.BackColor = System.Drawing.Color.LightYellow
        Me.LOGOUTToolStripMenuItem.Name = "LOGOUTToolStripMenuItem"
        Me.LOGOUTToolStripMenuItem.Size = New System.Drawing.Size(177, 34)
        Me.LOGOUTToolStripMenuItem.Text = "LOG OUT"
        '
        'ADMINToolStripMenuItem
        '
        Me.ADMINToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.POSSYSTEMToolStripMenuItem, Me.ProductInventoryToolStripMenuItem, Me.CustomerListToolStripMenuItem})
        Me.ADMINToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ADMINToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.ADMINToolStripMenuItem.Name = "ADMINToolStripMenuItem"
        Me.ADMINToolStripMenuItem.Size = New System.Drawing.Size(91, 34)
        Me.ADMINToolStripMenuItem.Text = "GO TO"
        '
        'POSSYSTEMToolStripMenuItem
        '
        Me.POSSYSTEMToolStripMenuItem.Name = "POSSYSTEMToolStripMenuItem"
        Me.POSSYSTEMToolStripMenuItem.Size = New System.Drawing.Size(259, 34)
        Me.POSSYSTEMToolStripMenuItem.Text = "POS SYSTEM"
        '
        'ProductInventoryToolStripMenuItem
        '
        Me.ProductInventoryToolStripMenuItem.Name = "ProductInventoryToolStripMenuItem"
        Me.ProductInventoryToolStripMenuItem.Size = New System.Drawing.Size(259, 34)
        Me.ProductInventoryToolStripMenuItem.Text = "Product Inventory"
        '
        'CustomerListToolStripMenuItem
        '
        Me.CustomerListToolStripMenuItem.Name = "CustomerListToolStripMenuItem"
        Me.CustomerListToolStripMenuItem.Size = New System.Drawing.Size(259, 34)
        Me.CustomerListToolStripMenuItem.Text = "Customer List"
        '
        'DASHBOARDToolStripMenuItem
        '
        Me.DASHBOARDToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DASHBOARDToolStripMenuItem.ForeColor = System.Drawing.Color.White
        Me.DASHBOARDToolStripMenuItem.Name = "DASHBOARDToolStripMenuItem"
        Me.DASHBOARDToolStripMenuItem.Size = New System.Drawing.Size(167, 34)
        Me.DASHBOARDToolStripMenuItem.Text = "SALES REPORT"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Courier New", 13.8!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(28, 46)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(582, 88)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "HI there! Welcome to Super Bites mini grocery store." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "We're happy to have you wit" & _
    "h us today. " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Feel free to look around, and if you need anything, " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "we're always" & _
    " here to help. Happy shopping!"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.OliveDrab
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Location = New System.Drawing.Point(407, 327)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(612, 186)
        Me.Panel3.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.LightYellow
        Me.Label3.Font = New System.Drawing.Font("Kristen ITC", 15.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.DarkOliveGreen
        Me.Label3.Location = New System.Drawing.Point(1252, 758)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(135, 28)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Super Bites"
        '
        'Form8
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Khaki
        Me.ClientSize = New System.Drawing.Size(1370, 749)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "Form8"
        Me.Text = "Form8"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents HOMEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GOTOToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SIGNUPToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LOGOUTToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ADMINToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DASHBOARDToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents POSSYSTEMToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProductInventoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CustomerListToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
