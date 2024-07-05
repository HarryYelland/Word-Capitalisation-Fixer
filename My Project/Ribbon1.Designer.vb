Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.Fixer = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.Tab1.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.Fixer.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Button1
        '
        Me.Button1.Label = "Button1"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Tab2
        '
        Me.Tab2.Groups.Add(Me.Fixer)
        Me.Tab2.Groups.Add(Me.Group2)
        Me.Tab2.Label = "Capitalisation Fixer"
        Me.Tab2.Name = "Tab2"
        '
        'Fixer
        '
        Me.Fixer.Items.Add(Me.Button2)
        Me.Fixer.Label = "Fixer"
        Me.Fixer.Name = "Fixer"
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Image = Global.CapitalisationFixerProject.My.Resources.Resources.Atoa1
        Me.Button2.Label = "Fix Highlighted Text"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Items.Add(Me.Label1)
        Me.Group2.Label = "Credits"
        Me.Group2.Name = "Group2"
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.Image = Global.CapitalisationFixerProject.My.Resources.Resources.Atoa2
        Me.Button3.Label = "By Harry Yelland"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Label1
        '
        Me.Label1.Label = "Created By Harry Yelland"
        Me.Label1.Name = "Label1"
        Me.Label1.ShowLabel = False
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        '
        'Ribbon1.OfficeMenu
        '
        Me.OfficeMenu.Items.Add(Me.Button1)
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Tab2.ResumeLayout(False)
        Me.Tab2.PerformLayout()
        Me.Fixer.ResumeLayout(False)
        Me.Fixer.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Fixer As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
