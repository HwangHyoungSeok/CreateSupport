<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreateSupportForm
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CreateSupportForm))
        lblTitle = New System.Windows.Forms.Label()
        lblClose = New System.Windows.Forms.Label()
        pnlTitle = New System.Windows.Forms.Panel()
        logoBox = New System.Windows.Forms.PictureBox()
        Preview = New System.Windows.Forms.PictureBox()
        nextbt = New System.Windows.Forms.Button()
        saddle = New System.Windows.Forms.CheckBox()
        grip = New System.Windows.Forms.CheckBox()
        btype = New System.Windows.Forms.CheckBox()
        ttext = New System.Windows.Forms.TextBox()
        Panel1 = New System.Windows.Forms.Panel()
        TextBox1 = New System.Windows.Forms.TextBox()
        TextBox3 = New System.Windows.Forms.TextBox()
        Panel2 = New System.Windows.Forms.Panel()
        mtext = New System.Windows.Forms.TextBox()
        Panel3 = New System.Windows.Forms.Panel()
        stext = New System.Windows.Forms.TextBox()
        TextBox4 = New System.Windows.Forms.TextBox()
        TextBox2 = New System.Windows.Forms.TextBox()
        TextBox5 = New System.Windows.Forms.TextBox()
        fselect = New System.Windows.Forms.Panel()
        PictureBox1 = New System.Windows.Forms.PictureBox()
        Label1 = New System.Windows.Forms.Label()
        length = New System.Windows.Forms.TextBox()
        Panel4 = New System.Windows.Forms.Panel()
        TextBox7 = New System.Windows.Forms.TextBox()
        TextBox8 = New System.Windows.Forms.TextBox()
        Panel5 = New System.Windows.Forms.Panel()
        Distance = New System.Windows.Forms.TextBox()
        TextBox10 = New System.Windows.Forms.TextBox()
        et1 = New System.Windows.Forms.CheckBox()
        et2 = New System.Windows.Forms.CheckBox()
        ok = New System.Windows.Forms.CheckBox()
        cancel = New System.Windows.Forms.CheckBox()
        apply = New System.Windows.Forms.CheckBox()
        Panel6 = New System.Windows.Forms.Panel()
        Model = New System.Windows.Forms.TextBox()
        TextBox11 = New System.Windows.Forms.TextBox()
        Panel7 = New System.Windows.Forms.Panel()
        TextBox14 = New System.Windows.Forms.TextBox()
        TextBox12 = New System.Windows.Forms.TextBox()
        TextBox13 = New System.Windows.Forms.TextBox()
        Panel8 = New System.Windows.Forms.Panel()
        splength = New System.Windows.Forms.TextBox()
        TextBox16 = New System.Windows.Forms.TextBox()
        Panel9 = New System.Windows.Forms.Panel()
        TextBox17 = New System.Windows.Forms.TextBox()
        TextBox18 = New System.Windows.Forms.TextBox()
        Panel10 = New System.Windows.Forms.Panel()
        filename = New System.Windows.Forms.TextBox()
        mbt = New System.Windows.Forms.CheckBox()
        sbt = New System.Windows.Forms.CheckBox()
        pnlTitle.SuspendLayout()
        CType(logoBox, ComponentModel.ISupportInitialize).BeginInit()
        CType(Preview, ComponentModel.ISupportInitialize).BeginInit()
        Panel1.SuspendLayout()
        Panel2.SuspendLayout()
        Panel3.SuspendLayout()
        fselect.SuspendLayout()
        CType(PictureBox1, ComponentModel.ISupportInitialize).BeginInit()
        Panel4.SuspendLayout()
        Panel5.SuspendLayout()
        Panel6.SuspendLayout()
        Panel7.SuspendLayout()
        Panel8.SuspendLayout()
        Panel9.SuspendLayout()
        Panel10.SuspendLayout()
        SuspendLayout()
        ' 
        ' lblTitle
        ' 
        lblTitle.Font = New System.Drawing.Font("맑은 고딕", 10F)
        lblTitle.ImageAlign = Drawing.ContentAlignment.TopCenter
        lblTitle.Location = New System.Drawing.Point(39, -1)
        lblTitle.Name = "lblTitle"
        lblTitle.Size = New System.Drawing.Size(736, 28)
        lblTitle.TabIndex = 5
        lblTitle.Text = "Creata Support"
        lblTitle.TextAlign = Drawing.ContentAlignment.MiddleLeft
        ' 
        ' lblClose
        ' 
        lblClose.Anchor = System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right
        lblClose.Font = New System.Drawing.Font("맑은 고딕", 11.25F)
        lblClose.Location = New System.Drawing.Point(781, 0)
        lblClose.Name = "lblClose"
        lblClose.Size = New System.Drawing.Size(28, 28)
        lblClose.TabIndex = 3
        lblClose.Text = "X"
        lblClose.TextAlign = Drawing.ContentAlignment.MiddleCenter
        ' 
        ' pnlTitle
        ' 
        pnlTitle.Controls.Add(lblClose)
        pnlTitle.Controls.Add(lblTitle)
        pnlTitle.Controls.Add(logoBox)
        pnlTitle.Dock = System.Windows.Forms.DockStyle.Top
        pnlTitle.Location = New System.Drawing.Point(0, 0)
        pnlTitle.Name = "pnlTitle"
        pnlTitle.Size = New System.Drawing.Size(809, 28)
        pnlTitle.TabIndex = 0
        ' 
        ' logoBox
        ' 
        logoBox.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        logoBox.Image = CType(resources.GetObject("logoBox.Image"), Drawing.Image)
        logoBox.Location = New System.Drawing.Point(2, -1)
        logoBox.Name = "logoBox"
        logoBox.Size = New System.Drawing.Size(34, 28)
        logoBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        logoBox.TabIndex = 14
        logoBox.TabStop = False
        ' 
        ' Preview
        ' 
        Preview.Location = New System.Drawing.Point(12, 35)
        Preview.Name = "Preview"
        Preview.Size = New System.Drawing.Size(209, 233)
        Preview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Preview.TabIndex = 1
        Preview.TabStop = False
        ' 
        ' nextbt
        ' 
        nextbt.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        nextbt.FlatAppearance.BorderSize = 0
        nextbt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        nextbt.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        nextbt.ForeColor = Drawing.Color.White
        nextbt.Location = New System.Drawing.Point(323, 87)
        nextbt.Name = "nextbt"
        nextbt.Size = New System.Drawing.Size(76, 27)
        nextbt.TabIndex = 10
        nextbt.Text = "NEXT"
        nextbt.UseVisualStyleBackColor = False
        ' 
        ' saddle
        ' 
        saddle.Appearance = System.Windows.Forms.Appearance.Button
        saddle.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        saddle.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        saddle.FlatAppearance.BorderSize = 0
        saddle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        saddle.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        saddle.ForeColor = Drawing.Color.White
        saddle.Location = New System.Drawing.Point(445, 116)
        saddle.Name = "saddle"
        saddle.Size = New System.Drawing.Size(76, 27)
        saddle.TabIndex = 11
        saddle.Text = "SADDLE"
        saddle.TextAlign = Drawing.ContentAlignment.MiddleCenter
        saddle.UseVisualStyleBackColor = False
        ' 
        ' grip
        ' 
        grip.Appearance = System.Windows.Forms.Appearance.Button
        grip.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        grip.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        grip.FlatAppearance.BorderSize = 0
        grip.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        grip.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        grip.ForeColor = Drawing.Color.White
        grip.Location = New System.Drawing.Point(527, 116)
        grip.Name = "grip"
        grip.Size = New System.Drawing.Size(76, 27)
        grip.TabIndex = 12
        grip.Text = "GRIP"
        grip.TextAlign = Drawing.ContentAlignment.MiddleCenter
        grip.UseVisualStyleBackColor = False
        ' 
        ' btype
        ' 
        btype.Appearance = System.Windows.Forms.Appearance.Button
        btype.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        btype.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        btype.FlatAppearance.BorderSize = 0
        btype.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        btype.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        btype.ForeColor = Drawing.Color.White
        btype.Location = New System.Drawing.Point(323, 120)
        btype.Name = "btype"
        btype.Size = New System.Drawing.Size(76, 27)
        btype.TabIndex = 13
        btype.Text = "Y/N"
        btype.TextAlign = Drawing.ContentAlignment.MiddleCenter
        btype.UseVisualStyleBackColor = False
        ' 
        ' ttext
        ' 
        ttext.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        ttext.BorderStyle = System.Windows.Forms.BorderStyle.None
        ttext.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        ttext.ForeColor = Drawing.Color.White
        ttext.Location = New System.Drawing.Point(-2, 7)
        ttext.Name = "ttext"
        ttext.Size = New System.Drawing.Size(85, 18)
        ttext.TabIndex = 15
        ttext.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' Panel1
        ' 
        Panel1.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel1.Controls.Add(ttext)
        Panel1.Location = New System.Drawing.Point(322, 54)
        Panel1.Name = "Panel1"
        Panel1.Size = New System.Drawing.Size(85, 27)
        Panel1.TabIndex = 16
        ' 
        ' TextBox1
        ' 
        TextBox1.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox1.ForeColor = Drawing.Color.White
        TextBox1.Location = New System.Drawing.Point(329, 37)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New System.Drawing.Size(69, 16)
        TextBox1.TabIndex = 17
        TextBox1.Text = "Category"
        TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' TextBox3
        ' 
        TextBox3.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox3.ForeColor = Drawing.Color.White
        TextBox3.Location = New System.Drawing.Point(239, 37)
        TextBox3.Name = "TextBox3"
        TextBox3.Size = New System.Drawing.Size(69, 16)
        TextBox3.TabIndex = 19
        TextBox3.Text = "Material"
        TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' Panel2
        ' 
        Panel2.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel2.Controls.Add(mtext)
        Panel2.Location = New System.Drawing.Point(232, 54)
        Panel2.Name = "Panel2"
        Panel2.Size = New System.Drawing.Size(85, 27)
        Panel2.TabIndex = 20
        ' 
        ' mtext
        ' 
        mtext.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        mtext.BorderStyle = System.Windows.Forms.BorderStyle.None
        mtext.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        mtext.ForeColor = Drawing.Color.White
        mtext.Location = New System.Drawing.Point(3, 8)
        mtext.Name = "mtext"
        mtext.Size = New System.Drawing.Size(85, 18)
        mtext.TabIndex = 15
        mtext.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' Panel3
        ' 
        Panel3.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel3.Controls.Add(stext)
        Panel3.Location = New System.Drawing.Point(501, 54)
        Panel3.Name = "Panel3"
        Panel3.Size = New System.Drawing.Size(85, 27)
        Panel3.TabIndex = 22
        ' 
        ' stext
        ' 
        stext.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        stext.BorderStyle = System.Windows.Forms.BorderStyle.None
        stext.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        stext.ForeColor = Drawing.Color.White
        stext.Location = New System.Drawing.Point(-2, 6)
        stext.Name = "stext"
        stext.Size = New System.Drawing.Size(85, 18)
        stext.TabIndex = 15
        stext.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' TextBox4
        ' 
        TextBox4.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox4.ForeColor = Drawing.Color.White
        TextBox4.Location = New System.Drawing.Point(508, 37)
        TextBox4.Name = "TextBox4"
        TextBox4.Size = New System.Drawing.Size(69, 16)
        TextBox4.TabIndex = 21
        TextBox4.Text = "SIZE"
        TextBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' TextBox2
        ' 
        TextBox2.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox2.Font = New System.Drawing.Font("맑은 고딕", 9.75F)
        TextBox2.ForeColor = Drawing.Color.White
        TextBox2.Location = New System.Drawing.Point(236, 92)
        TextBox2.Name = "TextBox2"
        TextBox2.Size = New System.Drawing.Size(76, 18)
        TextBox2.TabIndex = 23
        TextBox2.Text = "Foot Type"
        TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        ' 
        ' TextBox5
        ' 
        TextBox5.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox5.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox5.Font = New System.Drawing.Font("맑은 고딕", 9.75F)
        TextBox5.ForeColor = Drawing.Color.White
        TextBox5.Location = New System.Drawing.Point(236, 125)
        TextBox5.Name = "TextBox5"
        TextBox5.Size = New System.Drawing.Size(76, 18)
        TextBox5.TabIndex = 24
        TextBox5.Text = "Bolt Type"
        TextBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        ' 
        ' fselect
        ' 
        fselect.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        fselect.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        fselect.Controls.Add(PictureBox1)
        fselect.Controls.Add(Label1)
        fselect.Location = New System.Drawing.Point(236, 190)
        fselect.Name = "fselect"
        fselect.Size = New System.Drawing.Size(107, 27)
        fselect.TabIndex = 27
        ' 
        ' PictureBox1
        ' 
        PictureBox1.BackColor = Drawing.Color.Transparent
        PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), Drawing.Image)
        PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        PictureBox1.Location = New System.Drawing.Point(10, 1)
        PictureBox1.Name = "PictureBox1"
        PictureBox1.Size = New System.Drawing.Size(34, 27)
        PictureBox1.TabIndex = 28
        PictureBox1.TabStop = False
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        Label1.ForeColor = Drawing.Color.White
        Label1.Location = New System.Drawing.Point(46, 6)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(52, 17)
        Label1.TabIndex = 28
        Label1.Text = "면 선택"
        ' 
        ' length
        ' 
        length.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        length.BorderStyle = System.Windows.Forms.BorderStyle.None
        length.Cursor = System.Windows.Forms.Cursors.IBeam
        length.Font = New System.Drawing.Font("맑은 고딕", 11.25F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        length.ForeColor = Drawing.Color.White
        length.Location = New System.Drawing.Point(4, 5)
        length.Name = "length"
        length.Size = New System.Drawing.Size(87, 20)
        length.TabIndex = 15
        length.Text = "0.00"
        length.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        ' 
        ' Panel4
        ' 
        Panel4.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel4.Controls.Add(length)
        Panel4.Location = New System.Drawing.Point(353, 191)
        Panel4.Name = "Panel4"
        Panel4.Size = New System.Drawing.Size(95, 27)
        Panel4.TabIndex = 28
        ' 
        ' TextBox7
        ' 
        TextBox7.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox7.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox7.Font = New System.Drawing.Font("맑은 고딕", 9.75F)
        TextBox7.ForeColor = Drawing.Color.White
        TextBox7.Location = New System.Drawing.Point(447, 200)
        TextBox7.Name = "TextBox7"
        TextBox7.Size = New System.Drawing.Size(27, 18)
        TextBox7.TabIndex = 29
        TextBox7.Text = "mm"
        ' 
        ' TextBox8
        ' 
        TextBox8.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox8.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox8.Font = New System.Drawing.Font("맑은 고딕", 9.75F)
        TextBox8.ForeColor = Drawing.Color.White
        TextBox8.Location = New System.Drawing.Point(448, 250)
        TextBox8.Name = "TextBox8"
        TextBox8.Size = New System.Drawing.Size(27, 18)
        TextBox8.TabIndex = 31
        TextBox8.Text = "mm"
        ' 
        ' Panel5
        ' 
        Panel5.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel5.Controls.Add(Distance)
        Panel5.Location = New System.Drawing.Point(354, 241)
        Panel5.Name = "Panel5"
        Panel5.Size = New System.Drawing.Size(94, 27)
        Panel5.TabIndex = 30
        ' 
        ' Distance
        ' 
        Distance.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Distance.BorderStyle = System.Windows.Forms.BorderStyle.None
        Distance.Cursor = System.Windows.Forms.Cursors.IBeam
        Distance.Font = New System.Drawing.Font("맑은 고딕", 11.25F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        Distance.ForeColor = Drawing.Color.White
        Distance.Location = New System.Drawing.Point(4, 4)
        Distance.Name = "Distance"
        Distance.Size = New System.Drawing.Size(87, 20)
        Distance.TabIndex = 15
        Distance.Text = "0.000"
        Distance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        ' 
        ' TextBox10
        ' 
        TextBox10.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox10.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox10.ForeColor = Drawing.Color.White
        TextBox10.Location = New System.Drawing.Point(367, 223)
        TextBox10.Name = "TextBox10"
        TextBox10.Size = New System.Drawing.Size(69, 16)
        TextBox10.TabIndex = 32
        TextBox10.Text = "간격 띄우기"
        TextBox10.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' et1
        ' 
        et1.Appearance = System.Windows.Forms.Appearance.Button
        et1.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        et1.BackgroundImage = CType(resources.GetObject("et1.BackgroundImage"), Drawing.Image)
        et1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        et1.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        et1.FlatAppearance.BorderSize = 0
        et1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        et1.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        et1.ForeColor = Drawing.Color.White
        et1.Location = New System.Drawing.Point(237, 228)
        et1.Name = "et1"
        et1.Size = New System.Drawing.Size(50, 40)
        et1.TabIndex = 35
        et1.TextAlign = Drawing.ContentAlignment.MiddleCenter
        et1.UseVisualStyleBackColor = False
        ' 
        ' et2
        ' 
        et2.Appearance = System.Windows.Forms.Appearance.Button
        et2.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        et2.BackgroundImage = CType(resources.GetObject("et2.BackgroundImage"), Drawing.Image)
        et2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        et2.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        et2.FlatAppearance.BorderSize = 0
        et2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        et2.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        et2.ForeColor = Drawing.Color.White
        et2.Location = New System.Drawing.Point(294, 228)
        et2.Name = "et2"
        et2.Size = New System.Drawing.Size(50, 40)
        et2.TabIndex = 36
        et2.TextAlign = Drawing.ContentAlignment.MiddleCenter
        et2.UseVisualStyleBackColor = False
        ' 
        ' ok
        ' 
        ok.Appearance = System.Windows.Forms.Appearance.Button
        ok.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        ok.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        ok.FlatAppearance.BorderSize = 0
        ok.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        ok.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        ok.ForeColor = Drawing.Color.White
        ok.Location = New System.Drawing.Point(523, 241)
        ok.Name = "ok"
        ok.Size = New System.Drawing.Size(80, 27)
        ok.TabIndex = 37
        ok.Text = "확인"
        ok.TextAlign = Drawing.ContentAlignment.MiddleCenter
        ok.UseVisualStyleBackColor = False
        ' 
        ' cancel
        ' 
        cancel.Appearance = System.Windows.Forms.Appearance.Button
        cancel.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        cancel.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        cancel.FlatAppearance.BorderSize = 0
        cancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        cancel.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        cancel.ForeColor = Drawing.Color.White
        cancel.Location = New System.Drawing.Point(612, 241)
        cancel.Name = "cancel"
        cancel.Size = New System.Drawing.Size(80, 27)
        cancel.TabIndex = 38
        cancel.Text = "취소"
        cancel.TextAlign = Drawing.ContentAlignment.MiddleCenter
        cancel.UseVisualStyleBackColor = False
        ' 
        ' apply
        ' 
        apply.Appearance = System.Windows.Forms.Appearance.Button
        apply.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        apply.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        apply.FlatAppearance.BorderSize = 0
        apply.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        apply.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        apply.ForeColor = Drawing.Color.White
        apply.Location = New System.Drawing.Point(700, 241)
        apply.Name = "apply"
        apply.Size = New System.Drawing.Size(80, 27)
        apply.TabIndex = 39
        apply.Text = "적용"
        apply.TextAlign = Drawing.ContentAlignment.MiddleCenter
        apply.UseVisualStyleBackColor = False
        ' 
        ' Panel6
        ' 
        Panel6.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel6.Controls.Add(Model)
        Panel6.Location = New System.Drawing.Point(411, 54)
        Panel6.Name = "Panel6"
        Panel6.Size = New System.Drawing.Size(85, 27)
        Panel6.TabIndex = 40
        ' 
        ' Model
        ' 
        Model.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Model.BorderStyle = System.Windows.Forms.BorderStyle.None
        Model.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        Model.ForeColor = Drawing.Color.White
        Model.Location = New System.Drawing.Point(-2, 7)
        Model.Name = "Model"
        Model.Size = New System.Drawing.Size(85, 18)
        Model.TabIndex = 15
        Model.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' TextBox11
        ' 
        TextBox11.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox11.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox11.ForeColor = Drawing.Color.White
        TextBox11.Location = New System.Drawing.Point(418, 37)
        TextBox11.Name = "TextBox11"
        TextBox11.Size = New System.Drawing.Size(69, 16)
        TextBox11.TabIndex = 41
        TextBox11.Text = "Model No."
        TextBox11.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' Panel7
        ' 
        Panel7.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel7.Controls.Add(TextBox14)
        Panel7.Controls.Add(TextBox12)
        Panel7.Location = New System.Drawing.Point(590, 54)
        Panel7.Name = "Panel7"
        Panel7.Size = New System.Drawing.Size(72, 27)
        Panel7.TabIndex = 43
        ' 
        ' TextBox14
        ' 
        TextBox14.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        TextBox14.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox14.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        TextBox14.ForeColor = Drawing.Color.White
        TextBox14.Location = New System.Drawing.Point(40, 7)
        TextBox14.Name = "TextBox14"
        TextBox14.Size = New System.Drawing.Size(30, 18)
        TextBox14.TabIndex = 15
        TextBox14.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' TextBox12
        ' 
        TextBox12.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        TextBox12.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox12.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        TextBox12.ForeColor = Drawing.Color.White
        TextBox12.Location = New System.Drawing.Point(2, 7)
        TextBox12.Name = "TextBox12"
        TextBox12.Size = New System.Drawing.Size(30, 18)
        TextBox12.TabIndex = 15
        TextBox12.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' TextBox13
        ' 
        TextBox13.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox13.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox13.ForeColor = Drawing.Color.White
        TextBox13.Location = New System.Drawing.Point(592, 37)
        TextBox13.Name = "TextBox13"
        TextBox13.Size = New System.Drawing.Size(69, 16)
        TextBox13.TabIndex = 42
        TextBox13.Text = "Pitch x 수량"
        TextBox13.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' Panel8
        ' 
        Panel8.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel8.Controls.Add(splength)
        Panel8.Location = New System.Drawing.Point(667, 54)
        Panel8.Name = "Panel8"
        Panel8.Size = New System.Drawing.Size(60, 27)
        Panel8.TabIndex = 45
        ' 
        ' splength
        ' 
        splength.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        splength.BorderStyle = System.Windows.Forms.BorderStyle.None
        splength.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        splength.ForeColor = Drawing.Color.White
        splength.Location = New System.Drawing.Point(-1, 6)
        splength.Name = "splength"
        splength.Size = New System.Drawing.Size(60, 18)
        splength.TabIndex = 15
        splength.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' TextBox16
        ' 
        TextBox16.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox16.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox16.ForeColor = Drawing.Color.White
        TextBox16.Location = New System.Drawing.Point(668, 37)
        TextBox16.Name = "TextBox16"
        TextBox16.Size = New System.Drawing.Size(60, 16)
        TextBox16.TabIndex = 44
        TextBox16.Text = "길이(mm)"
        TextBox16.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' Panel9
        ' 
        Panel9.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel9.Controls.Add(TextBox17)
        Panel9.Location = New System.Drawing.Point(733, 54)
        Panel9.Name = "Panel9"
        Panel9.Size = New System.Drawing.Size(60, 27)
        Panel9.TabIndex = 47
        ' 
        ' TextBox17
        ' 
        TextBox17.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        TextBox17.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox17.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        TextBox17.ForeColor = Drawing.Color.White
        TextBox17.Location = New System.Drawing.Point(-1, 6)
        TextBox17.Name = "TextBox17"
        TextBox17.Size = New System.Drawing.Size(60, 18)
        TextBox17.TabIndex = 15
        TextBox17.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' TextBox18
        ' 
        TextBox18.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        TextBox18.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox18.ForeColor = Drawing.Color.White
        TextBox18.Location = New System.Drawing.Point(734, 37)
        TextBox18.Name = "TextBox18"
        TextBox18.Size = New System.Drawing.Size(60, 16)
        TextBox18.TabIndex = 46
        TextBox18.Text = "경사(deg)"
        TextBox18.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' Panel10
        ' 
        Panel10.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        Panel10.Controls.Add(filename)
        Panel10.Location = New System.Drawing.Point(480, 190)
        Panel10.Name = "Panel10"
        Panel10.Size = New System.Drawing.Size(300, 30)
        Panel10.TabIndex = 51
        ' 
        ' filename
        ' 
        filename.BackColor = Drawing.Color.FromArgb(CByte(49), CByte(58), CByte(71))
        filename.BorderStyle = System.Windows.Forms.BorderStyle.None
        filename.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        filename.ForeColor = Drawing.Color.White
        filename.Location = New System.Drawing.Point(2, 6)
        filename.Name = "filename"
        filename.Size = New System.Drawing.Size(298, 18)
        filename.TabIndex = 15
        filename.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' mbt
        ' 
        mbt.Appearance = System.Windows.Forms.Appearance.Button
        mbt.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        mbt.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        mbt.FlatAppearance.BorderSize = 0
        mbt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        mbt.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        mbt.ForeColor = Drawing.Color.White
        mbt.Location = New System.Drawing.Point(323, 153)
        mbt.Name = "mbt"
        mbt.Size = New System.Drawing.Size(76, 27)
        mbt.TabIndex = 52
        mbt.Text = "Multi"
        mbt.TextAlign = Drawing.ContentAlignment.MiddleCenter
        mbt.UseVisualStyleBackColor = False
        ' 
        ' sbt
        ' 
        sbt.Appearance = System.Windows.Forms.Appearance.Button
        sbt.BackColor = Drawing.Color.FromArgb(CByte(78), CByte(86), CByte(100))
        sbt.CheckAlign = Drawing.ContentAlignment.MiddleCenter
        sbt.FlatAppearance.BorderSize = 0
        sbt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        sbt.Font = New System.Drawing.Font("맑은 고딕", 9.75F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(129))
        sbt.ForeColor = Drawing.Color.White
        sbt.Location = New System.Drawing.Point(239, 153)
        sbt.Name = "sbt"
        sbt.Size = New System.Drawing.Size(76, 27)
        sbt.TabIndex = 53
        sbt.Text = "Single"
        sbt.TextAlign = Drawing.ContentAlignment.MiddleCenter
        sbt.UseVisualStyleBackColor = False
        ' 
        ' CreateSupportForm
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7F, 15F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        BackColor = Drawing.Color.FromArgb(CByte(59), CByte(67), CByte(83))
        BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        ClientSize = New System.Drawing.Size(809, 348)
        Controls.Add(sbt)
        Controls.Add(mbt)
        Controls.Add(Panel10)
        Controls.Add(Panel9)
        Controls.Add(TextBox18)
        Controls.Add(Panel8)
        Controls.Add(TextBox16)
        Controls.Add(Panel7)
        Controls.Add(TextBox13)
        Controls.Add(Panel6)
        Controls.Add(TextBox11)
        Controls.Add(apply)
        Controls.Add(cancel)
        Controls.Add(ok)
        Controls.Add(et2)
        Controls.Add(et1)
        Controls.Add(TextBox10)
        Controls.Add(TextBox8)
        Controls.Add(Panel5)
        Controls.Add(TextBox7)
        Controls.Add(Panel4)
        Controls.Add(fselect)
        Controls.Add(TextBox5)
        Controls.Add(TextBox2)
        Controls.Add(Panel3)
        Controls.Add(TextBox4)
        Controls.Add(Panel2)
        Controls.Add(TextBox3)
        Controls.Add(TextBox1)
        Controls.Add(Panel1)
        Controls.Add(btype)
        Controls.Add(grip)
        Controls.Add(saddle)
        Controls.Add(nextbt)
        Controls.Add(Preview)
        Controls.Add(pnlTitle)
        ForeColor = Drawing.Color.Gainsboro
        FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Name = "CreateSupportForm"
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Text = "CreateSupportForm"
        pnlTitle.ResumeLayout(False)
        CType(logoBox, ComponentModel.ISupportInitialize).EndInit()
        CType(Preview, ComponentModel.ISupportInitialize).EndInit()
        Panel1.ResumeLayout(False)
        Panel1.PerformLayout()
        Panel2.ResumeLayout(False)
        Panel2.PerformLayout()
        Panel3.ResumeLayout(False)
        Panel3.PerformLayout()
        fselect.ResumeLayout(False)
        fselect.PerformLayout()
        CType(PictureBox1, ComponentModel.ISupportInitialize).EndInit()
        Panel4.ResumeLayout(False)
        Panel4.PerformLayout()
        Panel5.ResumeLayout(False)
        Panel5.PerformLayout()
        Panel6.ResumeLayout(False)
        Panel6.PerformLayout()
        Panel7.ResumeLayout(False)
        Panel7.PerformLayout()
        Panel8.ResumeLayout(False)
        Panel8.PerformLayout()
        Panel9.ResumeLayout(False)
        Panel9.PerformLayout()
        Panel10.ResumeLayout(False)
        Panel10.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblClose As System.Windows.Forms.Label
    Friend WithEvents pnlTitle As System.Windows.Forms.Panel
    Friend WithEvents Preview As System.Windows.Forms.PictureBox
    Friend WithEvents nextbt As System.Windows.Forms.Button
    Friend WithEvents logoBox As System.Windows.Forms.PictureBox
    Friend WithEvents saddle As System.Windows.Forms.CheckBox
    Friend WithEvents grip As System.Windows.Forms.CheckBox
    Friend WithEvents btype As System.Windows.Forms.CheckBox
    Friend WithEvents ttext As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents mtext As System.Windows.Forms.TextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents stext As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents fselect As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents length As System.Windows.Forms.TextBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Distance As System.Windows.Forms.TextBox
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents et1 As System.Windows.Forms.CheckBox
    Friend WithEvents et2 As System.Windows.Forms.CheckBox
    Friend WithEvents ok As System.Windows.Forms.CheckBox
    Friend WithEvents cancel As System.Windows.Forms.CheckBox
    Friend WithEvents apply As System.Windows.Forms.CheckBox
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Model As System.Windows.Forms.TextBox
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox13 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox14 As System.Windows.Forms.TextBox
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents splength As System.Windows.Forms.TextBox
    Friend WithEvents TextBox16 As System.Windows.Forms.TextBox
    Friend WithEvents Panel9 As System.Windows.Forms.Panel
    Friend WithEvents TextBox17 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox18 As System.Windows.Forms.TextBox
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents filename As System.Windows.Forms.TextBox
    Friend WithEvents mbt As System.Windows.Forms.CheckBox
    Friend WithEvents sbt As System.Windows.Forms.CheckBox
End Class
