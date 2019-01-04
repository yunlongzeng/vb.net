Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
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

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.SplitButton1 = Me.Factory.CreateRibbonSplitButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.SplitButton2 = Me.Factory.CreateRibbonSplitButton
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.SplitButton3 = Me.Factory.CreateRibbonSplitButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.SplitButton4 = Me.Factory.CreateRibbonSplitButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.SplitButton5 = Me.Factory.CreateRibbonSplitButton
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.SplitButton6 = Me.Factory.CreateRibbonSplitButton
        Me.Button22 = Me.Factory.CreateRibbonButton
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.EditBox1 = Me.Factory.CreateRibbonEditBox
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Button26 = Me.Factory.CreateRibbonButton
        Me.Button27 = Me.Factory.CreateRibbonButton
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.EditBox2 = Me.Factory.CreateRibbonEditBox
        Me.Button28 = Me.Factory.CreateRibbonButton
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.Button29 = Me.Factory.CreateRibbonButton
        Me.Group9 = Me.Factory.CreateRibbonGroup
        Me.Button34 = Me.Factory.CreateRibbonButton
        Me.Button30 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.Group7.SuspendLayout()
        Me.Group9.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Groups.Add(Me.Group7)
        Me.Tab1.Groups.Add(Me.Group9)
        Me.Tab1.Label = "CPQ ExcelView"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Button2)
        Me.Group1.Label = "Hide blanks"
        Me.Group1.Name = "Group1"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "HideBlanks"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "Unhide"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Label = "Ins Drawing"
        Me.Group2.Name = "Group2"
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Label = "GetInsDrawing"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.SplitButton1)
        Me.Group3.Label = "AddInstruments"
        Me.Group3.Name = "Group3"
        '
        'SplitButton1
        '
        Me.SplitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SplitButton1.Image = CType(resources.GetObject("SplitButton1.Image"), System.Drawing.Image)
        Me.SplitButton1.Items.Add(Me.Button4)
        Me.SplitButton1.Items.Add(Me.Button5)
        Me.SplitButton1.Items.Add(Me.SplitButton2)
        Me.SplitButton1.Items.Add(Me.Button6)
        Me.SplitButton1.Items.Add(Me.SplitButton3)
        Me.SplitButton1.Items.Add(Me.Button11)
        Me.SplitButton1.Items.Add(Me.Button12)
        Me.SplitButton1.Items.Add(Me.SplitButton4)
        Me.SplitButton1.Items.Add(Me.Button15)
        Me.SplitButton1.Items.Add(Me.Button19)
        Me.SplitButton1.Items.Add(Me.Button18)
        Me.SplitButton1.Items.Add(Me.SplitButton5)
        Me.SplitButton1.Items.Add(Me.Button17)
        Me.SplitButton1.Items.Add(Me.Button16)
        Me.SplitButton1.Items.Add(Me.SplitButton6)
        Me.SplitButton1.Label = "AddInstruments"
        Me.SplitButton1.Name = "SplitButton1"
        '
        'Button4
        '
        Me.Button4.Label = "AddPosition"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Button5
        '
        Me.Button5.Label = "AddLimitSwitch"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'SplitButton2
        '
        Me.SplitButton2.Items.Add(Me.Button7)
        Me.SplitButton2.Items.Add(Me.Button8)
        Me.SplitButton2.Label = "AddSOV"
        Me.SplitButton2.Name = "SplitButton2"
        '
        'Button7
        '
        Me.Button7.Label = "SOV1"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        '
        'Button8
        '
        Me.Button8.Label = "SOV2"
        Me.Button8.Name = "Button8"
        Me.Button8.ShowImage = True
        '
        'Button6
        '
        Me.Button6.Label = "AddFixingPlate"
        Me.Button6.Name = "Button6"
        Me.Button6.ShowImage = True
        '
        'SplitButton3
        '
        Me.SplitButton3.Items.Add(Me.Button9)
        Me.SplitButton3.Items.Add(Me.Button10)
        Me.SplitButton3.Label = "AddAOV"
        Me.SplitButton3.Name = "SplitButton3"
        '
        'Button9
        '
        Me.Button9.Label = "AOV1"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Button10
        '
        Me.Button10.Label = "AOV2"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
        '
        'Button11
        '
        Me.Button11.Label = "AddPPS"
        Me.Button11.Name = "Button11"
        Me.Button11.ShowImage = True
        '
        'Button12
        '
        Me.Button12.Label = "AddAFR"
        Me.Button12.Name = "Button12"
        Me.Button12.ShowImage = True
        '
        'SplitButton4
        '
        Me.SplitButton4.Items.Add(Me.Button13)
        Me.SplitButton4.Items.Add(Me.Button14)
        Me.SplitButton4.Label = "AddPressureGauge"
        Me.SplitButton4.Name = "SplitButton4"
        '
        'Button13
        '
        Me.Button13.Label = "PG1"
        Me.Button13.Name = "Button13"
        Me.Button13.ShowImage = True
        '
        'Button14
        '
        Me.Button14.Label = "PG2"
        Me.Button14.Name = "Button14"
        Me.Button14.ShowImage = True
        '
        'Button15
        '
        Me.Button15.Label = "AddVB"
        Me.Button15.Name = "Button15"
        Me.Button15.ShowImage = True
        '
        'Button19
        '
        Me.Button19.Label = "AddQEV"
        Me.Button19.Name = "Button19"
        Me.Button19.ShowImage = True
        '
        'Button18
        '
        Me.Button18.Label = "AddNeedleValve"
        Me.Button18.Name = "Button18"
        Me.Button18.ShowImage = True
        '
        'SplitButton5
        '
        Me.SplitButton5.Items.Add(Me.Button20)
        Me.SplitButton5.Items.Add(Me.Button21)
        Me.SplitButton5.Label = "AddCheckValve"
        Me.SplitButton5.Name = "SplitButton5"
        '
        'Button20
        '
        Me.Button20.Label = "CV1"
        Me.Button20.Name = "Button20"
        Me.Button20.ShowImage = True
        '
        'Button21
        '
        Me.Button21.Label = "CV2"
        Me.Button21.Name = "Button21"
        Me.Button21.ShowImage = True
        '
        'Button17
        '
        Me.Button17.Label = "AddSCV"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        '
        'Button16
        '
        Me.Button16.Label = "AddSilencer"
        Me.Button16.Name = "Button16"
        Me.Button16.ShowImage = True
        '
        'SplitButton6
        '
        Me.SplitButton6.Items.Add(Me.Button22)
        Me.SplitButton6.Items.Add(Me.Button23)
        Me.SplitButton6.Label = "AddOthers"
        Me.SplitButton6.Name = "SplitButton6"
        '
        'Button22
        '
        Me.Button22.Label = "Other1"
        Me.Button22.Name = "Button22"
        Me.Button22.ShowImage = True
        '
        'Button23
        '
        Me.Button23.Label = "Other2"
        Me.Button23.Name = "Button23"
        Me.Button23.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button24)
        Me.Group4.Items.Add(Me.EditBox1)
        Me.Group4.Items.Add(Me.Button25)
        Me.Group4.Label = "WaitingToCheck"
        Me.Group4.Name = "Group4"
        '
        'Button24
        '
        Me.Button24.Image = CType(resources.GetObject("Button24.Image"), System.Drawing.Image)
        Me.Button24.Label = "SaveToWTC"
        Me.Button24.Name = "Button24"
        Me.Button24.ShowImage = True
        '
        'EditBox1
        '
        Me.EditBox1.Label = "CPQNo."
        Me.EditBox1.Name = "EditBox1"
        Me.EditBox1.Text = Nothing
        '
        'Button25
        '
        Me.Button25.Image = CType(resources.GetObject("Button25.Image"), System.Drawing.Image)
        Me.Button25.Label = "OpenWTCExcel"
        Me.Button25.Name = "Button25"
        Me.Button25.ShowImage = True
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.Button26)
        Me.Group5.Items.Add(Me.Button27)
        Me.Group5.Label = "DoneExcel"
        Me.Group5.Name = "Group5"
        '
        'Button26
        '
        Me.Button26.Image = CType(resources.GetObject("Button26.Image"), System.Drawing.Image)
        Me.Button26.Label = "MoveToDone"
        Me.Button26.Name = "Button26"
        Me.Button26.ShowImage = True
        '
        'Button27
        '
        Me.Button27.Image = CType(resources.GetObject("Button27.Image"), System.Drawing.Image)
        Me.Button27.Label = "OpenDoneExcel"
        Me.Button27.Name = "Button27"
        Me.Button27.ShowImage = True
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.EditBox2)
        Me.Group6.Items.Add(Me.Button28)
        Me.Group6.Label = "SaveExcelToInput_Excel"
        Me.Group6.Name = "Group6"
        '
        'EditBox2
        '
        Me.EditBox2.Label = "CO"
        Me.EditBox2.Name = "EditBox2"
        Me.EditBox2.Text = Nothing
        '
        'Button28
        '
        Me.Button28.Image = CType(resources.GetObject("Button28.Image"), System.Drawing.Image)
        Me.Button28.Label = "SaveToInputFolder"
        Me.Button28.Name = "Button28"
        Me.Button28.ShowImage = True
        '
        'Group7
        '
        Me.Group7.DialogLauncher = RibbonDialogLauncherImpl1
        Me.Group7.Items.Add(Me.Button29)
        Me.Group7.Label = "Accessories"
        Me.Group7.Name = "Group7"
        '
        'Button29
        '
        Me.Button29.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button29.Image = CType(resources.GetObject("Button29.Image"), System.Drawing.Image)
        Me.Button29.Label = "Create Accessoreis"
        Me.Button29.Name = "Button29"
        Me.Button29.ShowImage = True
        '
        'Group9
        '
        Me.Group9.Items.Add(Me.Button34)
        Me.Group9.Items.Add(Me.Button30)
        Me.Group9.Label = "IQI"
        Me.Group9.Name = "Group9"
        '
        'Button34
        '
        Me.Button34.Image = CType(resources.GetObject("Button34.Image"), System.Drawing.Image)
        Me.Button34.Label = "Comment"
        Me.Button34.Name = "Button34"
        Me.Button34.ShowImage = True
        '
        'Button30
        '
        Me.Button30.Image = CType(resources.GetObject("Button30.Image"), System.Drawing.Image)
        Me.Button30.Label = "Sort"
        Me.Button30.Name = "Button30"
        Me.Button30.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
        Me.Group9.ResumeLayout(False)
        Me.Group9.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SplitButton1 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton2 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton3 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton4 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton5 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton6 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditBox1 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button26 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button27 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button28 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditBox2 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button29 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button30 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group9 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button34 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
