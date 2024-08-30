<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class index
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(index))
        Me.AxWebBrowser1 = New AxSHDocVw.AxWebBrowser()
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxWebBrowser1
        '
        Me.AxWebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AxWebBrowser1.Enabled = True
        Me.AxWebBrowser1.Location = New System.Drawing.Point(0, 0)
        Me.AxWebBrowser1.OcxState = CType(resources.GetObject("AxWebBrowser1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxWebBrowser1.Size = New System.Drawing.Size(909, 622)
        Me.AxWebBrowser1.TabIndex = 0
        '
        'index
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(909, 622)
        Me.Controls.Add(Me.AxWebBrowser1)
        Me.Name = "index"
        Me.Tag = "index"
        Me.Text = "WEB"
        CType(Me.AxWebBrowser1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents AxWebBrowser1 As AxSHDocVw.AxWebBrowser
End Class
