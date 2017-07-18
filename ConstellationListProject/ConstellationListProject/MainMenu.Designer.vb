<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenu
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.MainMenuBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.MainMenuBindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataSet1 = New System.Data.DataSet()
        Me.DataSet2 = New System.Data.DataSet()
        Me.DataSet3 = New System.Data.DataSet()
        Me.MainMenuBindingSource2 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.MainMenuBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainMenuBindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainMenuBindingSource2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("メイリオ", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(81, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(172, 55)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "星座表示"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(31, 282)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 33)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "星座検索"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(222, 282)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(84, 33)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "マスタメンテ"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'MainMenuBindingSource
        '
        Me.MainMenuBindingSource.DataSource = GetType(ConstellationListProject.MainMenu)
        '
        'MainMenuBindingSource1
        '
        Me.MainMenuBindingSource1.DataSource = GetType(ConstellationListProject.MainMenu)
        '
        'DataSet1
        '
        Me.DataSet1.DataSetName = "NewDataSet"
        '
        'DataSet2
        '
        Me.DataSet2.DataSetName = "NewDataSet"
        '
        'DataSet3
        '
        Me.DataSet3.DataSetName = "NewDataSet"
        '
        'MainMenuBindingSource2
        '
        Me.MainMenuBindingSource2.DataSource = GetType(ConstellationListProject.MainMenu)
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(222, 236)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(45, 23)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "＞＞"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(91, 236)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(41, 23)
        Me.Button4.TabIndex = 7
        Me.Button4.Text = "＜＜"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(91, 93)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(176, 137)
        Me.TextBox1.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(155, 241)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(0, 12)
        Me.Label2.TabIndex = 9
        '
        'MainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(345, 328)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "MainMenu"
        Me.Text = "MainMenu"
        CType(Me.MainMenuBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainMenuBindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainMenuBindingSource2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents MainMenuBindingSource As BindingSource
    Friend WithEvents MainMenuBindingSource1 As BindingSource
    Friend WithEvents DataSet1 As DataSet
    Friend WithEvents DataSet2 As DataSet
    Friend WithEvents DataSet3 As DataSet
    Friend WithEvents MainMenuBindingSource2 As BindingSource
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label2 As Label
End Class
