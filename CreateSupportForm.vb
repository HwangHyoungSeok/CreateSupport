' ========================= CreateSupportForm.vb =========================
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Inventor
Imports Draw = System.Drawing
Imports IO = System.IO

Public Class CreateSupportForm

#Region "A. 공통/Win32 + 버튼 기본 동작(옵션/토글/다음) "

    ' [NEW] Apply 후 UI/상태를 "초기 상태"로 되돌림
    Private Sub ResetUIToInitialState()
        Try
            ' 텍스트 박스 초기화
            If ttext IsNot Nothing Then ttext.Text = ""
            If mtext IsNot Nothing Then mtext.Text = ""
            If stext IsNot Nothing Then stext.Text = ""
            If Model IsNot Nothing Then Model.Text = ""
            If length IsNot Nothing Then length.Text = ""
            If splength IsNot Nothing Then splength.Text = ""
            If filename IsNot Nothing Then filename.Text = ""

            ' Distance 기본값
            If Distance IsNot Nothing Then Distance.Text = "0.00"

            ' 옵션/토글 버튼 초기화
            ResetAllButtons()
            UpdateOptionButtons()
            InitSingleMultiToggle()

        Catch
        End Try
    End Sub

    ' [NEW] 창 드래그 시작
    Private Sub DragArea_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Left Then Return
        isDragging = True
        dragStartPoint = e.Location
    End Sub

    ' [NEW] 창 드래그 이동
    Private Sub DragArea_MouseMove(sender As Object, e As MouseEventArgs)
        If Not isDragging Then Return
        Dim p = Me.PointToScreen(e.Location)
        Me.Location = New Draw.Point(p.X - dragStartPoint.X, p.Y - dragStartPoint.Y)
    End Sub

    ' [NEW] 창 드래그 종료
    Private Sub DragArea_MouseUp(sender As Object, e As MouseEventArgs)
        isDragging = False
    End Sub

    ' --- LegType 순환(nextbt) ---
    Private Sub nextbt_Click(sender As Object, e As EventArgs) Handles nextbt.Click

        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

        ' ----------------------------
        ' PVC: CF <-> GF 순환 (기본 CF)
        ' ----------------------------
        If mat = "PVC" Then
            If m_legCodePVC = "CF" Then
                m_legCodePVC = "GF"
            Else
                m_legCodePVC = "CF"
            End If

            UpdateFilename()

            If m_supportFace IsNot Nothing Then
                Me.BeginInvoke(New Action(AddressOf RunPreviewPlacement))
            End If
            Return
        End If


        ' ----------------------------
        ' STS: 기존 로직 그대로 (SF->SS->SR->SL 순환)
        ' ----------------------------
        If mat <> "STS" Then
            UpdateFilename()
            Return
        End If

        Select Case m_legCode
            Case "SF" : m_legCode = "SS"
            Case "SS" : m_legCode = "SR"
            Case "SR" : m_legCode = "SL"
            Case Else : m_legCode = "SF"
        End Select

        UpdateFilename()

        If m_supportFace IsNot Nothing Then
            Me.BeginInvoke(New Action(AddressOf RunPreviewPlacement))
        End If
    End Sub


    ' --- Win32 ---
    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    ' --- 공통 상태 ---
    Private isFormRunning As Boolean = False
    Private isDragging As Boolean = False
    Private dragStartPoint As Draw.Point

    ' --- 버튼 리셋/옵션 ---
    Private Sub ResetAllButtons()
        saddle.Checked = False
        grip.Checked = False
        et1.Checked = False
        et2.Checked = False
        ok.Checked = False
        cancel.Checked = False
        apply.Checked = False

        If sbt IsNot Nothing Then sbt.Checked = True
        If mbt IsNot Nothing Then mbt.Checked = False
    End Sub

    Private Sub UpdateOptionButtons()
        Dim enableMain As Boolean = Not (saddle.Checked OrElse grip.Checked)

        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

        ' nextbt는 PVC도 써야 하니까 enableMain 유지
        nextbt.Enabled = enableMain

        ' btype은 PVC면 항상 잠금
        If mat = "PVC" Then
            btype.Enabled = False
            btype.Checked = False
            btype.BackColor = Draw.Color.FromArgb(60, 60, 60)
        Else
            btype.Enabled = enableMain
        End If


        If enableMain Then
            btype.BackColor = cBtnNormal
            nextbt.BackColor = cBtnNormal
        Else
            btype.Checked = False
            btype.BackColor = Draw.Color.FromArgb(60, 60, 60)
            nextbt.BackColor = Draw.Color.FromArgb(60, 60, 60)
        End If

    End Sub

    Private Sub btype_CheckedChanged(sender As Object, e As EventArgs) Handles btype.CheckedChanged
        UpdateFilename()

        ' 이미 면 선택까지 끝나서 미리보기 떠있으면 즉시 갱신
        If m_supportFace IsNot Nothing Then
            Me.BeginInvoke(New Action(AddressOf RunPreviewPlacement))
        End If
    End Sub

    Private Sub SaveTempCreatedFiles()
        Try
            For Each p In m_tempCreatedFiles
                If IO.File.Exists(p) Then
                    Try
                        Dim d As Document = Globals.g_inventorApplication.Documents.Open(p, False)
                        d.Update2(True)
                        d.Save2(True)
                        d.Close(True)
                    Catch
                    End Try
                End If
            Next
        Catch
        End Try
    End Sub

#End Region

#Region "B. Single/Multi 토글(sbt/mbt) + 호버/클릭 색상"

    Private isUpdatingSM As Boolean = False

    Private Sub InitSingleMultiToggle()
        ' 기본값: Single ON
        isUpdatingSM = True
        Try
            If sbt IsNot Nothing Then
                sbt.Appearance = Appearance.Button
                sbt.FlatStyle = FlatStyle.Flat
                sbt.FlatAppearance.BorderSize = 0
                sbt.Checked = True
                sbt.BackColor = cBtnNormal
            End If

            If mbt IsNot Nothing Then
                mbt.Appearance = Appearance.Button
                mbt.FlatStyle = FlatStyle.Flat
                mbt.FlatAppearance.BorderSize = 0
                mbt.Checked = False
                mbt.BackColor = cBtnNormal
            End If
        Finally
            isUpdatingSM = False
        End Try

        UpdateFilename()

        ' 이벤트 연결(중복 방지)
        If sbt IsNot Nothing Then
            RemoveHandler sbt.CheckedChanged, AddressOf sbt_CheckedChanged
            AddHandler sbt.CheckedChanged, AddressOf sbt_CheckedChanged

            RemoveHandler sbt.MouseEnter, AddressOf SM_MouseEnter
            RemoveHandler sbt.MouseLeave, AddressOf SM_MouseLeave
            RemoveHandler sbt.MouseDown, AddressOf SM_MouseDown
            RemoveHandler sbt.MouseUp, AddressOf SM_MouseUp

            AddHandler sbt.MouseEnter, AddressOf SM_MouseEnter
            AddHandler sbt.MouseLeave, AddressOf SM_MouseLeave
            AddHandler sbt.MouseDown, AddressOf SM_MouseDown
            AddHandler sbt.MouseUp, AddressOf SM_MouseUp
        End If

        If mbt IsNot Nothing Then
            RemoveHandler mbt.CheckedChanged, AddressOf mbt_CheckedChanged
            AddHandler mbt.CheckedChanged, AddressOf mbt_CheckedChanged

            RemoveHandler mbt.MouseEnter, AddressOf SM_MouseEnter
            RemoveHandler mbt.MouseLeave, AddressOf SM_MouseLeave
            RemoveHandler mbt.MouseDown, AddressOf SM_MouseDown
            RemoveHandler mbt.MouseUp, AddressOf SM_MouseUp

            AddHandler mbt.MouseEnter, AddressOf SM_MouseEnter
            AddHandler mbt.MouseLeave, AddressOf SM_MouseLeave
            AddHandler mbt.MouseDown, AddressOf SM_MouseDown
            AddHandler mbt.MouseUp, AddressOf SM_MouseUp
        End If
    End Sub

    Private Sub sbt_CheckedChanged(sender As Object, e As EventArgs)
        If isUpdatingSM Then Return
        If sbt Is Nothing OrElse mbt Is Nothing Then Return

        isUpdatingSM = True
        Try
            If sbt.Checked Then
                mbt.Checked = False
            Else
                ' 항상 하나는 ON
                sbt.Checked = True
            End If
        Finally
            isUpdatingSM = False
        End Try
        UpdateFilename()
    End Sub

    Private Sub mbt_CheckedChanged(sender As Object, e As EventArgs)
        If isUpdatingSM Then Return
        If sbt Is Nothing OrElse mbt Is Nothing Then Return

        isUpdatingSM = True
        Try
            If mbt.Checked Then
                sbt.Checked = False
            Else
                mbt.Checked = True
            End If
        Finally
            isUpdatingSM = False
        End Try
        UpdateFilename()
    End Sub

    ' 호버/클릭 색상: 기존 버튼과 동일 컨셉
    Private Sub SM_MouseEnter(sender As Object, e As EventArgs)
        Dim cb = TryCast(sender, CheckBox)
        If cb Is Nothing Then Return
        cb.BackColor = cBtnHover
    End Sub

    Private Sub SM_MouseLeave(sender As Object, e As EventArgs)
        Dim cb = TryCast(sender, CheckBox)
        If cb Is Nothing Then Return
        cb.BackColor = cBtnNormal
    End Sub

    Private Sub SM_MouseDown(sender As Object, e As MouseEventArgs)
        Dim cb = TryCast(sender, CheckBox)
        If cb Is Nothing Then Return
        If e.Button = MouseButtons.Left Then cb.BackColor = cBtnClick
    End Sub

    Private Sub SM_MouseUp(sender As Object, e As MouseEventArgs)
        Dim cb = TryCast(sender, CheckBox)
        If cb Is Nothing Then Return
        cb.BackColor = cBtnHover
    End Sub

#End Region

#Region "C. 옵션 체크(Saddle/Grip) + UI 색상"

    Private Sub saddle_CheckedChanged(sender As Object, e As EventArgs) Handles saddle.CheckedChanged
        If saddle.Checked Then grip.Checked = False
        UpdateOptionButtons()
    End Sub

    Private Sub grip_CheckedChanged(sender As Object, e As EventArgs) Handles grip.CheckedChanged
        If grip.Checked Then saddle.Checked = False
        UpdateOptionButtons()
    End Sub

    Private ReadOnly cGlyph As Draw.Color = Draw.Color.FromArgb(180, 190, 200)
    Private ReadOnly cCloseHover As Draw.Color = Draw.Color.FromArgb(235, 120, 120)

    Private ReadOnly cBtnNormal As Draw.Color = Draw.Color.FromArgb(78, 86, 100)
    Private ReadOnly cBtnHover As Draw.Color = Draw.Color.FromArgb(95, 105, 120)
    Private ReadOnly cBtnClick As Draw.Color = Draw.Color.FromArgb(60, 68, 80)

#End Region

#Region "D. Inventor 선택/하이라이트 상태(변수) + Mate/Flush 변수"
    ' =========================
    ' [NEW] 예외 무시 안전 실행 유틸
    ' =========================
    Private Sub SafeRun(act As Action)
        Try
            act()
        Catch
        End Try
    End Sub


    ' [NEW] 축 슬라이드/날아감 방지용: 축방향 위치 잠금(평면 구속)
    Private m_tempAxialLockPlane As WorkPlane = Nothing
    Private m_tempAxialLockConstraint As AssemblyConstraint = Nothing

    ' [NEW] 현재 미리보기로 쓰고 있는 Support(어셈블리 폴더 ipt) 경로
    Private m_currentSupportAsmPath As String = ""

    ' [NEW] 현재 미리보기(임시배치)로 사용 중인 파일 경로(어셈블리 폴더 안 ipt)
    Private m_currentPreviewSupportAsmPath As String = ""
    Private m_currentPreviewSadAsmPath As String = ""

    ' [NEW] 사용자가 모델을 클릭한 3D 지점(어셈블리 좌표)
    Private m_lastPickPoint As Inventor.Point = Nothing

    ' [NEW] 임시로 복사/생성한 파일 목록(취소/X 시 삭제, Apply/OK 시 유지)
    Private m_tempCreatedFiles As New List(Of String)

    ' [NEW] 현재 배치된 SAD의 소스 경로 기억(같으면 재배치/Replace 안 함)
    Private m_currentSadSourcePath As String = ""

    Private m_blueHighlight As HighlightSet
    Private m_greenHighlight As HighlightSet

    ' 축대축(중심축) 전용 임시 구속
    Private m_tempAxisConstraint As AssemblyConstraint = Nothing

    ' Mate/Flush 전용 임시 구속
    Private m_tempMateFlushConstraint As AssemblyConstraint = Nothing

    Private m_tempWorkAxis As WorkAxis = Nothing

    ' [NEW] 스폰 위치 고정용(축 미끄럼 방지)
    Private m_tempSpawnWorkPoint As WorkPoint = Nothing
    Private m_tempSpawnConstraint As AssemblyConstraint = Nothing


    ' Mate/Flush 면 하이라이트(확정 전까지 유지)
    Private m_hlBaseFace As HighlightSet = Nothing
    Private m_hlTargetFace As HighlightSet = Nothing

    Private m_baseFace As Face
    Private m_targetFace As Face
    Private m_isFlushMode As Boolean = False

    Private isObjectSelectionActive As Boolean = False

    Private m_selectedComponent As ComponentOccurrence
    Private m_centerPoint As Inventor.Point
    Private m_supportFace As Object   ' Face 또는 FaceProxy

    Private m_supportCompHighlight As HighlightSet
    Private m_supportFaceHighlight As HighlightSet

    ' STS LegType 코드 (기본 SF)
    Private m_legCode As String = "SF"

    ' PVC FootType 코드 (기본 CF)  CF <-> GF
    Private m_legCodePVC As String = "CF"


    ' sbt/mbt 토글 중 재귀 방지(현재 코드에 존재)
    Private m_isToggleUpdating As Boolean = False

    ' [NEW] 찾은 support 파일 경로(Apply에서 사용)
    Private m_resolvedSupportPath As String = ""

    ' [NEW] 어셈블리에 배치된 Support Occurrence(다음 단계 구속에 사용)
    Private m_supportOcc As ComponentOccurrence = Nothing

    ' [NEW] SAD 파일 임시 배치 Occurrence
    Private m_sadOcc As ComponentOccurrence = Nothing

    ' [NEW] 축대축(중심축) 임시 구속 2개 (Support용 / SAD용)
    Private m_tempAxisConstraint_Support As AssemblyConstraint = Nothing
    Private m_tempAxisConstraint_SAD As AssemblyConstraint = Nothing

    ' [NEW] 찾은 SAD 파일 경로
    Private m_resolvedSadPath As String = ""

#End Region

#Region "E. SUPPORT 선택 헬퍼(SetSupportComponent / CenterPoint)"

    Private Sub SetSupportComponent(comp As ComponentOccurrence)
        m_selectedComponent = comp
        m_centerPoint = GetSupportCenterPoint(comp)

        ' 선택된 컴포넌트 파란 하이라이트
        Try
            Dim doc = Globals.g_inventorApplication.ActiveDocument
            Dim tobj = Globals.g_inventorApplication.TransientObjects

            If m_supportCompHighlight Is Nothing Then
                m_supportCompHighlight = doc.CreateHighlightSet()
                m_supportCompHighlight.Color = tobj.CreateColor(0, 191, 255)
            End If

            m_supportCompHighlight.Clear()
            m_supportCompHighlight.AddItem(comp)
            Globals.g_inventorApplication.ActiveView.Update()
        Catch
        End Try
    End Sub

    Private Function GetSupportCenterPoint(comp As ComponentOccurrence) As Inventor.Point
        Try
            ' 1) 원통면이 있으면 그 원통 BasePoint를 중심으로 사용
            For Each body As SurfaceBody In comp.Definition.SurfaceBodies
                For Each f As Face In body.Faces
                    If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        Dim cyl As Cylinder = CType(f.Geometry, Cylinder)
                        Dim pt As Inventor.Point = cyl.BasePoint
                        pt.TransformBy(comp.Transformation) ' 전역 좌표로 변환
                        Return pt
                    End If
                Next
            Next

            ' 2) 원통면 없으면 RangeBox 중앙
            Dim minP = comp.RangeBox.MinPoint
            Dim maxP = comp.RangeBox.MaxPoint

            Dim x As Double = (minP.X + maxP.X) / 2
            Dim y As Double = (minP.Y + maxP.Y) / 2
            Dim z As Double = (minP.Z + maxP.Z) / 2

            Return Globals.g_inventorApplication.TransientGeometry.CreatePoint(x, y, z)
        Catch
            Return Nothing
        End Try
    End Function

#End Region

#Region "F. Preview 상태(변수만)"
    ' [NEW] SAD는 "인치(튜브)"면 TU로 찾고, "A(파이프)"면 PI로 찾는다
    Private Function GetSadTypeCode() As String
        Dim s As String = If(stext.Text, "").Trim()

        If s.Contains("""") Then
            Return "TU"   ' 1/2", 3/4", 1" ...
        End If

        If s.ToUpper().EndsWith("A") Then
            Return "PI"   ' 15A, 20A, 25A ...
        End If

        ' 애매하면 기존 Category 기반
        Return ConvertTypeToCode(ttext.Text)
    End Function


#End Region

#Region "G. Form LifeCycle(Load/Shown/Close) + Cancel"

    Private Sub CreateSupportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DoubleBuffered = True
        Me.KeyPreview = True
        isFormRunning = True

        Dim buttons As CheckBox() = {ok, cancel, apply, et1, et2, sbt, mbt}
        For Each btn As CheckBox In buttons
            If btn IsNot Nothing Then
                btn.FlatStyle = FlatStyle.Flat
                btn.FlatAppearance.BorderSize = 0
                btn.BackColor = cBtnNormal
            End If
        Next

        If lblClose IsNot Nothing Then
            lblClose.ForeColor = cGlyph
            lblClose.Cursor = Cursors.Hand
            lblClose.BringToFront()

            ' 중복 연결 방지 후 1회만 연결
            RemoveHandler lblClose.Click, AddressOf lblClose_Click
            RemoveHandler lblClose.MouseEnter, AddressOf lblClose_MouseEnter
            RemoveHandler lblClose.MouseLeave, AddressOf lblClose_MouseLeave

            AddHandler lblClose.Click, AddressOf lblClose_Click
            AddHandler lblClose.MouseEnter, AddressOf lblClose_MouseEnter
            AddHandler lblClose.MouseLeave, AddressOf lblClose_MouseLeave
        End If

        AddHandler Distance.MouseWheel, AddressOf Distance_MouseWheel
        AddHandler Distance.TextChanged, AddressOf Distance_TextChanged

        ' [NEW] 드래그로 창 이동
        AddHandler Me.MouseDown, AddressOf DragArea_MouseDown
        AddHandler Me.MouseMove, AddressOf DragArea_MouseMove
        AddHandler Me.MouseUp, AddressOf DragArea_MouseUp

        If pnlTitle IsNot Nothing Then
            AddHandler pnlTitle.MouseDown, AddressOf DragArea_MouseDown
            AddHandler pnlTitle.MouseMove, AddressOf DragArea_MouseMove
            AddHandler pnlTitle.MouseUp, AddressOf DragArea_MouseUp
        End If

        If lblTitle IsNot Nothing Then
            AddHandler lblTitle.MouseDown, AddressOf DragArea_MouseDown
            AddHandler lblTitle.MouseMove, AddressOf DragArea_MouseMove
            AddHandler lblTitle.MouseUp, AddressOf DragArea_MouseUp
        End If

        If logoBox IsNot Nothing Then
            AddHandler logoBox.MouseDown, AddressOf DragArea_MouseDown
            AddHandler logoBox.MouseMove, AddressOf DragArea_MouseMove
            AddHandler logoBox.MouseUp, AddressOf DragArea_MouseUp
        End If

        ResetAllButtons()
        UpdateOptionButtons()
        InitSingleMultiToggle()
        ApplyMaterialRules()   ' ★ 추가
        UpdateFilename()

    End Sub

    Private Sub CreateSupportForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Try
            System.Windows.Forms.Application.DoEvents()
            RunAnalysisSequence()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "RunAnalysisSequence 오류")
        End Try
    End Sub

    Private Sub CreateSupportForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        isFormRunning = False
        Try
            If Globals.g_inventorApplication IsNot Nothing Then
                SetForegroundWindow(New IntPtr(Globals.g_inventorApplication.MainFrameHWND))
                SendKeys.SendWait("{ESC}")
            End If
        Catch
        End Try
        ClearAll_OnClose()
    End Sub

    ' ===================== [NEW] 임시 생성 파일 삭제 유틸 =====================
    Private Sub DeleteTempCreatedFiles()
        Try
            ' ★ Inventor가 문서를 열어둔 상태면 File.Delete가 실패한다.
            '   따라서 삭제 전에 열려있으면 닫고, 그래도 안 되면 조용히 스킵한다.
            For Each p In m_tempCreatedFiles
                If String.IsNullOrWhiteSpace(p) Then Continue For

                If IO.File.Exists(p) Then
                    ' 열려있으면 닫기(잠김 해제)
                    TryCloseDocumentIfOpen(p)

                    ' 그래도 잠겨있을 수 있으니 Try로 보호
                    Try
                        IO.File.Delete(p)
                    Catch
                    End Try
                End If
            Next
        Catch
        End Try

        m_tempCreatedFiles.Clear()
    End Sub

    ' [NEW] 파일이 Inventor에서 열려있으면 닫아서 삭제 가능하게 함
    Private Sub TryCloseDocumentIfOpen(fullPath As String)
        Try
            If Globals.g_inventorApplication Is Nothing Then Return
            If String.IsNullOrWhiteSpace(fullPath) Then Return

            Dim docs = Globals.g_inventorApplication.Documents
            If docs Is Nothing Then Return

            For i As Integer = docs.Count To 1 Step -1
                Dim d As Document = Nothing
                Try
                    d = docs.Item(i)
                Catch
                    Continue For
                End Try
                If d Is Nothing Then Continue For

                Dim p As String = ""
                Try
                    p = d.FullFileName
                Catch
                    Continue For
                End Try
                If String.IsNullOrWhiteSpace(p) Then Continue For

                If String.Equals(p, fullPath, StringComparison.OrdinalIgnoreCase) Then
                    Try
                        d.Close(True)
                    Catch
                    End Try
                    Exit For
                End If

            Next
        Catch
        End Try
    End Sub


    ' [NEW] Cancel/X/닫기 시: 임시배치(occ) + 임시구속 + 임시축 + 하이라이트까지 전부 정리
    Private Sub ClearAll_OnClose()

        Try
            If m_supportOcc IsNot Nothing Then m_supportOcc.Delete()
        Catch
        End Try
        m_supportOcc = Nothing

        Try
            If m_sadOcc IsNot Nothing Then m_sadOcc.Delete()
        Catch
        End Try
        m_sadOcc = Nothing

        Try
            If m_tempAxisConstraint_Support IsNot Nothing Then m_tempAxisConstraint_Support.Delete()
        Catch
        End Try
        m_tempAxisConstraint_Support = Nothing

        Try
            If m_tempAxisConstraint_SAD IsNot Nothing Then m_tempAxisConstraint_SAD.Delete()
        Catch
        End Try
        m_tempAxisConstraint_SAD = Nothing

        Try
            If m_tempMateFlushConstraint IsNot Nothing Then m_tempMateFlushConstraint.Delete()
        Catch
        End Try
        m_tempMateFlushConstraint = Nothing

        Try
            If m_tempWorkAxis IsNot Nothing Then m_tempWorkAxis.Delete()
        Catch
        End Try
        m_tempWorkAxis = Nothing

        ' [NEW] 스폰/축방향 잠금용 구속/WorkPoint/WorkPlane 삭제
        Try
            If m_tempSpawnConstraint IsNot Nothing Then m_tempSpawnConstraint.Delete()
        Catch
        End Try
        m_tempSpawnConstraint = Nothing

        Try
            If m_tempSpawnWorkPoint IsNot Nothing Then m_tempSpawnWorkPoint.Delete()
        Catch
        End Try
        m_tempSpawnWorkPoint = Nothing

        Try
            If m_tempAxialLockConstraint IsNot Nothing Then m_tempAxialLockConstraint.Delete()
        Catch
        End Try
        m_tempAxialLockConstraint = Nothing

        Try
            If m_tempAxialLockPlane IsNot Nothing Then m_tempAxialLockPlane.Delete()
        Catch
        End Try
        m_tempAxialLockPlane = Nothing



        ClearMateFlushFaceHighlights()

        Try
            If m_supportCompHighlight IsNot Nothing Then m_supportCompHighlight.Delete()
        Catch
        End Try
        m_supportCompHighlight = Nothing

        Try
            If m_supportFaceHighlight IsNot Nothing Then m_supportFaceHighlight.Delete()
        Catch
        End Try
        m_supportFaceHighlight = Nothing

        m_baseFace = Nothing
        m_targetFace = Nothing
        m_selectedComponent = Nothing
        m_centerPoint = Nothing
        m_supportFace = Nothing

        DeleteTempCreatedFiles()
    End Sub

    Private Sub cancel_Click(sender As Object, e As EventArgs) Handles cancel.Click
        ClearAll_OnClose()
        Me.Close()
    End Sub

    ' [REPLACE] Apply: 배치 확정(Commit) + UI 초기화 + 바로 다음 TUBE/PIPE 선택으로 진행
    Private Sub apply_Click(sender As Object, e As EventArgs) Handles apply.Click
        If m_supportOcc Is Nothing Then
            MessageBox.Show("배치된 Support가 없습니다." & vbCrLf & "먼저 '면 선택'을 진행하여 Support를 배치하세요.", "알림")
            Return
        End If

        ' 1) 확정(Commit)
        '   - 현재 미리보기 TMP를 final로 승격/합류한 뒤,
        '   - Occurrence는 final로 Replace (확정 후 공유는 여기서만 허용)
        If m_supportOcc IsNot Nothing Then

            ' ★ OK 확정 직전: 현재 TMP(메모리)의 파라미터를 디스크에 저장
            Try
                Dim partDef As PartComponentDefinition = TryCast(m_supportOcc.Definition, PartComponentDefinition)
                Dim pdoc As PartDocument = TryCast(partDef.Document, PartDocument)
                If pdoc IsNot Nothing Then
                    Try : pdoc.Update2(True) : Catch : End Try
                    Try : pdoc.Save2(True) : Catch : End Try
                End If
            Catch
            End Try

            Dim finalFileName As String = If(filename.Text, "").Trim()
            If finalFileName <> "" AndAlso m_currentPreviewSupportAsmPath <> "" Then

                Dim finalPath As String = CommitPreviewToFinal(m_currentPreviewSupportAsmPath, finalFileName)

                If finalPath <> "" Then
                    Try
                        m_supportOcc.Replace(finalPath, True)
                        Globals.g_inventorApplication.ActiveView.Update()
                    Catch
                    End Try
                End If
            End If
        End If


        SafeRun(Sub()
                    If m_tempAxisConstraint_Support IsNot Nothing Then m_tempAxisConstraint_Support.Delete()
                End Sub)

        SafeRun(Sub()
                    If m_tempAxisConstraint_SAD IsNot Nothing Then m_tempAxisConstraint_SAD.Delete()
                End Sub)

        SafeRun(Sub()
                    If m_tempMateFlushConstraint IsNot Nothing Then m_tempMateFlushConstraint.Delete()
                End Sub)

        SafeRun(Sub()
                    If m_tempSpawnConstraint IsNot Nothing Then m_tempSpawnConstraint.Delete()
                End Sub)

        SafeRun(Sub()
                    If m_tempSpawnWorkPoint IsNot Nothing Then m_tempSpawnWorkPoint.Delete()
                End Sub)

        SafeRun(Sub()
                    If m_tempAxialLockConstraint IsNot Nothing Then m_tempAxialLockConstraint.Delete()
                End Sub)

        SafeRun(Sub()
                    If m_tempAxialLockPlane IsNot Nothing Then m_tempAxialLockPlane.Delete()
                End Sub)


        m_tempAxisConstraint_Support = Nothing
        m_tempAxisConstraint_SAD = Nothing
        m_tempMateFlushConstraint = Nothing
        m_tempSpawnConstraint = Nothing
        m_tempSpawnWorkPoint = Nothing
        m_tempAxialLockConstraint = Nothing
        m_tempAxialLockPlane = Nothing


        m_supportOcc = Nothing
        m_sadOcc = Nothing
        m_currentSadSourcePath = ""
        m_currentSupportAsmPath = ""
        m_currentPreviewSupportAsmPath = ""

        ClearMateFlushFaceHighlights()

        ' Apply에서만 저장
        SaveTempCreatedFiles()
        m_tempCreatedFiles.Clear()


        ' 2) 다음 선택 준비
        Try
            If Globals.g_inventorApplication IsNot Nothing Then
                SetForegroundWindow(New IntPtr(Globals.g_inventorApplication.MainFrameHWND))
                SendKeys.SendWait("{ESC}")
            End If
        Catch
        End Try

        ' 3) 내부 선택 상태 초기화
        m_selectedComponent = Nothing
        m_centerPoint = Nothing
        m_supportFace = Nothing
        m_baseFace = Nothing
        m_targetFace = Nothing

        ' 4) UI 초기 상태 리셋
        ResetUIToInitialState()

        ' 5) 바로 다음 선택 루프로 복귀
        isObjectSelectionActive = True
        Me.BeginInvoke(New Action(AddressOf RunAnalysisSequence))
    End Sub

    ' [NEW] OK: 임시 배치를 확정하고 창 닫기
    Private Sub ok_Click(sender As Object, e As EventArgs) Handles ok.Click

        ' OK도 Apply처럼 확정(Commit) 수행 후 닫기
        If m_supportOcc IsNot Nothing Then

            Dim finalFileName As String = If(filename.Text, "").Trim()
            If finalFileName <> "" AndAlso m_currentPreviewSupportAsmPath <> "" Then

                Dim finalPath As String = CommitPreviewToFinal(m_currentPreviewSupportAsmPath, finalFileName)

                If finalPath <> "" Then
                    Try
                        m_supportOcc.Replace(finalPath, True)
                        Globals.g_inventorApplication.ActiveView.Update()
                    Catch
                    End Try
                End If
            End If
        End If

        If m_supportOcc IsNot Nothing OrElse m_sadOcc IsNot Nothing Then
            m_tempAxisConstraint_Support = Nothing
            m_tempAxisConstraint_SAD = Nothing
            m_tempMateFlushConstraint = Nothing
            m_tempSpawnConstraint = Nothing
            m_tempSpawnWorkPoint = Nothing

            m_supportOcc = Nothing
            m_sadOcc = Nothing
            m_currentSadSourcePath = ""
            m_currentPreviewSupportAsmPath = ""
            m_currentSupportAsmPath = ""
        End If

        ClearMateFlushFaceHighlights()
        SaveTempCreatedFiles()
        m_tempCreatedFiles.Clear()

        Me.Close()
    End Sub

    Private Sub RunPreviewPlacement()
        Try

            ' ★ mat 먼저 선언 (이게 핵심)
            Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

            m_resolvedSupportPath = FindSupportFile()
            If String.IsNullOrWhiteSpace(m_resolvedSupportPath) Then
                MessageBox.Show(
        "조건에 맞는 Support 파일을 찾지 못했습니다." & vbCrLf &
        "폴더: " & GetSupportFolderPath() & vbCrLf &
        "키: " & mat & " / " & ConvertSingleMultiToCode() & " / " & GetLegCode() & " / " & ConvertBoltToCode(),
        "찾기 실패"
    )
                Return
            End If

            m_resolvedSadPath = ""
            If mat = "STS" Then
                Dim stsFolder As String = IO.Path.Combine(SUPPORT_BASE_PATH, "STS_SUPPORT")
                m_resolvedSadPath = FindStsSadFile(stsFolder)
                If String.IsNullOrWhiteSpace(m_resolvedSadPath) Then
                    m_resolvedSadPath = ""
                End If
            Else
                ' PVC 포함: SAD 없음
                m_resolvedSadPath = ""
            End If


            Dim newFileName As String = If(filename.Text, "").Trim()
            If String.IsNullOrWhiteSpace(newFileName) Then Return

            Dim asm As AssemblyDocument = TryCast(Globals.g_inventorApplication.ActiveDocument, AssemblyDocument)
            If asm Is Nothing Then Return

            ' ==============================
            ' Support: [CHANGED] 미리보기는 항상 TMP만 사용 (final 재사용 금지)
            '   - 옵션 변경/취소는 TMP만 건드림
            '   - final 합류/재사용은 Apply/OK에서만 허용
            ' ==============================
            Dim prevPreviewPath As String = m_currentPreviewSupportAsmPath

            ' 미리보기 TMP 생성(항상 유니크)
            Dim previewPath As String = CreatePreviewTempSupport(m_resolvedSupportPath, newFileName)
            If String.IsNullOrWhiteSpace(previewPath) Then Return

            ' 이번 세션 임시 생성 목록에 등록(취소/X 시 삭제 대상)
            If Not m_tempCreatedFiles.Contains(previewPath) Then
                m_tempCreatedFiles.Add(previewPath)
            End If

            If m_supportOcc Is Nothing Then
                Dim spawnPt As Inventor.Point = GetSpawnPointOnTargetAxis()
                m_supportOcc = PlaceSupportOccurrence(previewPath, spawnPt)
                If m_supportOcc Is Nothing Then Return
            Else
                Try
                    m_supportOcc.Replace(previewPath, True)
                    Globals.g_inventorApplication.ActiveView.Update()
                Catch
                    Try : m_supportOcc.Delete() : Catch : End Try
                    Dim spawnPt As Inventor.Point = GetSpawnPointOnTargetAxis()
                    m_supportOcc = PlaceSupportOccurrence(previewPath, spawnPt)
                    If m_supportOcc Is Nothing Then Return
                End Try
            End If


            ' 현재 미리보기 TMP 경로만 갱신
            m_currentPreviewSupportAsmPath = previewPath

            ' 이전 미리보기 TMP 정리(세션 임시 생성 파일만 삭제)
            CleanupPrevPreviewTempFile(prevPreviewPath, m_currentPreviewSupportAsmPath)


            ' 파라미터 갱신(항상 수행)
            Dim sdVal As Double = GetSizeOffset(stext.Text)

            Dim rawLenText As String = If(length.Text, "").Trim()
            Dim splVal As Double = 0
            If Not Double.TryParse(rawLenText,
                               Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture,
                               splVal) Then
                Double.TryParse(rawLenText, splVal)
            End If

            Dim matU As String = mat
            Dim setSD As Boolean = (matU <> "PVC")   ' PVC는 SD 수정 안 함
            UpdateSupportParameters(m_supportOcc, sdVal, splVal, setSD, True)


            ' ==============================
            ' SAD 배치(선택) - STS일 때만
            ' ==============================
            If mat = "STS" AndAlso Not String.IsNullOrWhiteSpace(m_resolvedSadPath) Then

                Dim desiredSadPath As String = m_resolvedSadPath.Trim()

                If m_sadOcc IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(m_currentSadSourcePath) AndAlso
               String.Equals(m_currentSadSourcePath, desiredSadPath, StringComparison.OrdinalIgnoreCase) Then
                    ' 그대로 유지
                Else
                    If m_sadOcc Is Nothing Then
                        Dim spawnPtSad As Inventor.Point = GetSpawnPointOnTargetAxis()
                        m_sadOcc = PlaceSupportOccurrence(desiredSadPath, spawnPtSad)
                        If m_sadOcc Is Nothing Then Return
                    Else
                        Try
                            m_sadOcc.Replace(desiredSadPath, True)
                            Globals.g_inventorApplication.ActiveView.Update()
                        Catch
                            Try : m_sadOcc.Delete() : Catch : End Try
                            Dim spawnPtSad As Inventor.Point = GetSpawnPointOnTargetAxis()
                            m_sadOcc = PlaceSupportOccurrence(desiredSadPath, spawnPtSad)
                            If m_sadOcc Is Nothing Then Return
                        End Try
                    End If

                    m_currentSadSourcePath = desiredSadPath
                End If
            End If

            ' ==============================
            ' 축대축 임시구속
            ' ==============================
            If m_selectedComponent Is Nothing Then Return

            ' matU는 위에서 이미 선언돼 있으니 재선언하지 말고 그대로 사용
            Dim preferZ As Boolean = (matU = "PVC")

            Dim spawnPtForLock As Inventor.Point = GetSpawnPointOnTargetAxis()

            If ApplyTempAxisToAxisPreview(m_supportOcc, m_selectedComponent, m_tempAxisConstraint_Support, preferZ, spawnPtForLock) Then
                Globals.g_inventorApplication.StatusBarText = "Support(기본) 미리보기 배치 완료"
            End If

            If mat = "STS" AndAlso m_sadOcc IsNot Nothing Then
                If ApplyTempAxisToAxisPreview(m_sadOcc, m_selectedComponent, m_tempAxisConstraint_SAD, True, spawnPtForLock) Then
                    Globals.g_inventorApplication.StatusBarText = "Support(기본) + SAD 미리보기 배치 완료"
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "RunPreviewPlacement 오류")
        End Try
    End Sub

    ' =========================
    ' [NEW] 미리보기 전용: 무조건 TMP 복사본 생성
    ' =========================
    Private Function CreatePreviewTempSupport(srcLibPath As String, finalFileName As String) As String
        Dim asm As AssemblyDocument = TryCast(Globals.g_inventorApplication.ActiveDocument, AssemblyDocument)
        If asm Is Nothing Then Return ""
        If String.IsNullOrWhiteSpace(srcLibPath) OrElse Not IO.File.Exists(srcLibPath) Then Return ""
        If String.IsNullOrWhiteSpace(finalFileName) Then Return ""

        finalFileName = finalFileName.Replace(".", "P")

        Dim asmFolder As String = IO.Path.GetDirectoryName(asm.FullFileName)

        ' ★ TMP는 항상 유니크
        Dim tmpName As String = finalFileName & "__TMP_" & DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") & "_" & Guid.NewGuid().ToString("N").Substring(0, 6)
        Dim tmpPath As String = IO.Path.Combine(asmFolder, tmpName & ".ipt")

        Try
            IO.File.Copy(srcLibPath, tmpPath, False)
            Try : IO.File.SetAttributes(tmpPath, IO.FileAttributes.Normal) : Catch : End Try
            Return tmpPath
        Catch
            Return ""
        End Try
    End Function


    ' =========================
    ' [CHANGED] 확정(Apply/OK):
    '   - final이 이미 있으면 TMP 내용을 final에 "덮어쓰기(merge)" 후 TMP 삭제
    '   - final이 없으면 TMP를 final로 Move(승격)
    ' =========================
    Private Function CommitPreviewToFinal(tmpPath As String, finalFileName As String) As String
        Dim asm As AssemblyDocument = TryCast(Globals.g_inventorApplication.ActiveDocument, AssemblyDocument)
        If asm Is Nothing Then Return ""
        If String.IsNullOrWhiteSpace(finalFileName) Then Return ""

        finalFileName = finalFileName.Replace(".", "P")

        Dim asmFolder As String = IO.Path.GetDirectoryName(asm.FullFileName)
        Dim finalPath As String = IO.Path.Combine(asmFolder, finalFileName & ".ipt")

        If String.IsNullOrWhiteSpace(tmpPath) Then Return ""
        If Not IO.File.Exists(tmpPath) Then Return ""

        ' 1) final이 이미 있으면: TMP -> final 덮어쓰기(파라미터 포함) + TMP 삭제
        If IO.File.Exists(finalPath) Then
            Try
                TryCloseDocumentIfOpen(finalPath)
                TryCloseDocumentIfOpen(tmpPath)

                Try : IO.File.SetAttributes(finalPath, IO.FileAttributes.Normal) : Catch : End Try
                Try : IO.File.SetAttributes(tmpPath, IO.FileAttributes.Normal) : Catch : End Try

                IO.File.Copy(tmpPath, finalPath, True)

                For i As Integer = 1 To 5
                    Try
                        IO.File.Delete(tmpPath)
                        Exit For
                    Catch
                        Try : System.Threading.Thread.Sleep(150) : Catch : End Try
                    End Try
                Next

                Return finalPath
            Catch
                Return finalPath
            End Try
        End If

        ' 2) final이 없으면: TMP를 final로 승격
        Try
            TryCloseDocumentIfOpen(tmpPath)
            IO.File.Move(tmpPath, finalPath)
            Try : IO.File.SetAttributes(finalPath, IO.FileAttributes.Normal) : Catch : End Try
            Return finalPath
        Catch
        End Try

        Return ""
    End Function




    ' [NEW] et1=Mate / et2=Flush
    Private Sub et1_CheckedChanged(sender As Object, e As EventArgs) Handles et1.CheckedChanged
        If et1 Is Nothing Then Return
        If Not et1.Checked Then Return

        m_isFlushMode = False
        If et2 IsNot Nothing Then et2.Checked = False

        If m_baseFace IsNot Nothing AndAlso m_targetFace IsNot Nothing Then
            RebuildTempMateFlushConstraint()
        Else
            Me.BeginInvoke(New Action(AddressOf StartMateFlushPickSequence))
        End If
    End Sub

    Private Sub et2_CheckedChanged(sender As Object, e As EventArgs) Handles et2.CheckedChanged
        If et2 Is Nothing Then Return
        If Not et2.Checked Then Return

        m_isFlushMode = True
        If et1 IsNot Nothing Then et1.Checked = False

        If m_baseFace IsNot Nothing AndAlso m_targetFace IsNot Nothing Then
            RebuildTempMateFlushConstraint()
        Else
            Me.BeginInvoke(New Action(AddressOf StartMateFlushPickSequence))
        End If
    End Sub

    ' [NEW] Mate/Flush 면선택 → 임시 구속 생성
    Private Sub StartMateFlushPickSequence()
        If Globals.g_inventorApplication Is Nothing Then Return

        Try
            If m_tempMateFlushConstraint IsNot Nothing Then m_tempMateFlushConstraint.Delete()
        Catch
        End Try
        m_tempMateFlushConstraint = Nothing

        m_baseFace = Nothing
        m_targetFace = Nothing

        Try
            SetForegroundWindow(New IntPtr(Globals.g_inventorApplication.MainFrameHWND))
            SendKeys.SendWait("{ESC}")
        Catch
        End Try

        Try
            Dim f1 = Globals.g_inventorApplication.CommandManager.Pick(
                SelectionFilterEnum.kPartFaceFilter,
                "기준 면(Base Face)을 선택하세요")

            If f1 Is Nothing OrElse Not TypeOf f1 Is Face Then Return
            m_baseFace = CType(f1, Face)
            HighlightMateFlushFace(m_baseFace, True)

            Dim f2 = Globals.g_inventorApplication.CommandManager.Pick(
                SelectionFilterEnum.kPartFaceFilter,
                "맞출 면(Target Face)을 선택하세요")

            If f2 Is Nothing OrElse Not TypeOf f2 Is Face Then Return
            m_targetFace = CType(f2, Face)
            HighlightMateFlushFace(m_targetFace, False)

            RebuildTempMateFlushConstraint()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Mate/Flush 선택 오류")
        End Try
    End Sub

    ' [NEW] 현재 모드(Mate/Flush) + Distance 값으로 임시구속 생성/갱신
    Private Sub RebuildTempMateFlushConstraint()
        If Globals.g_inventorApplication Is Nothing Then Return
        If m_baseFace Is Nothing OrElse m_targetFace Is Nothing Then Return

        Dim doc As Document = Globals.g_inventorApplication.ActiveDocument
        If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Return

        Dim asm As AssemblyDocument = CType(doc, AssemblyDocument)
        Dim asmDef As AssemblyComponentDefinition = asm.ComponentDefinition
        Dim uom As UnitsOfMeasure = asm.UnitsOfMeasure

        Try
            If m_tempMateFlushConstraint IsNot Nothing Then m_tempMateFlushConstraint.Delete()
        Catch
        End Try
        m_tempMateFlushConstraint = Nothing

        Dim offsetDb As Double = GetOffsetDbFromDistanceText(uom)

        Try
            If m_isFlushMode Then
                m_tempMateFlushConstraint = asmDef.Constraints.AddFlushConstraint(m_baseFace, m_targetFace, offsetDb)
            Else
                m_tempMateFlushConstraint = asmDef.Constraints.AddMateConstraint(m_baseFace, m_targetFace, offsetDb)
            End If

            Globals.g_inventorApplication.ActiveView.Update()
        Catch ex As Exception
            MessageBox.Show("임시 구속 생성 실패: " & ex.Message, "구속 오류")
        End Try
    End Sub

    ' [NEW] Distance 텍스트 변경 시 실시간 업데이트
    Private Sub Distance_TextChanged(sender As Object, e As EventArgs)
        If Globals.g_inventorApplication Is Nothing Then Return
        If m_tempMateFlushConstraint Is Nothing Then Return

        Dim doc As Document = Globals.g_inventorApplication.ActiveDocument
        Dim asm As AssemblyDocument = TryCast(doc, AssemblyDocument)
        If asm Is Nothing Then Return

        Try
            Dim uom As UnitsOfMeasure = asm.UnitsOfMeasure
            Dim offsetDb As Double = GetOffsetDbFromDistanceText(uom)

            Dim mc As MateConstraint = TryCast(m_tempMateFlushConstraint, MateConstraint)
            If mc IsNot Nothing Then
                mc.Offset.Value = offsetDb
            Else
                Dim fc As FlushConstraint = TryCast(m_tempMateFlushConstraint, FlushConstraint)
                If fc IsNot Nothing Then fc.Offset.Value = offsetDb
            End If

            asm.Update2(True)
            Globals.g_inventorApplication.ActiveView.Update()
        Catch
        End Try
    End Sub

    ' [NEW] Distance.Text(mm) → DB length 변환
    Private Function GetOffsetDbFromDistanceText(uom As UnitsOfMeasure) As Double
        Dim s As String = If(Distance.Text, "").Trim()
        If s = "" Then Return 0

        Dim v As Double = 0
        If Not Double.TryParse(s, Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture, v) Then
            Double.TryParse(s, v)
        End If

        Return uom.GetValueFromExpression(
            v.ToString("0.###", Globalization.CultureInfo.InvariantCulture) & " mm",
            UnitsTypeEnum.kMillimeterLengthUnits
        )
    End Function

    Private Sub ClearMateFlushFaceHighlights()
        Try
            If m_hlBaseFace IsNot Nothing Then m_hlBaseFace.Delete()
        Catch
        End Try
        m_hlBaseFace = Nothing

        Try
            If m_hlTargetFace IsNot Nothing Then m_hlTargetFace.Delete()
        Catch
        End Try
        m_hlTargetFace = Nothing

        Try
            If Globals.g_inventorApplication IsNot Nothing Then
                Globals.g_inventorApplication.ActiveView.Update()
            End If
        Catch
        End Try
    End Sub

    Private Sub HighlightMateFlushFace(faceObj As Face, isBase As Boolean)
        Try
            Dim app = Globals.g_inventorApplication
            If app Is Nothing OrElse faceObj Is Nothing Then Return

            Dim doc = app.ActiveDocument
            If doc Is Nothing Then Return

            Dim hs As HighlightSet = Nothing

            If isBase Then
                If m_hlBaseFace Is Nothing Then
                    m_hlBaseFace = doc.CreateHighlightSet()
                    m_hlBaseFace.Color = app.TransientObjects.CreateColor(0, 191, 255)
                End If
                hs = m_hlBaseFace
            Else
                If m_hlTargetFace Is Nothing Then
                    m_hlTargetFace = doc.CreateHighlightSet()
                    m_hlTargetFace.Color = app.TransientObjects.CreateColor(120, 255, 120)
                End If
                hs = m_hlTargetFace
            End If

            hs.Clear()
            hs.AddItem(faceObj)
            app.ActiveView.Update()
        Catch
        End Try
    End Sub

    ' [NEW] X 클릭 = 취소와 동일
    Private Sub lblClose_Click(sender As Object, e As EventArgs)
        ClearAll_OnClose()
        Me.Close()
    End Sub

    Private Sub lblClose_MouseEnter(sender As Object, e As EventArgs)
        If lblClose Is Nothing Then Return
        lblClose.ForeColor = cCloseHover
    End Sub

    Private Sub lblClose_MouseLeave(sender As Object, e As EventArgs)
        If lblClose Is Nothing Then Return
        lblClose.ForeColor = cGlyph
    End Sub

#End Region

#Region "H. 면선택(fselect) 버튼 이벤트(Click/Hover)"

    Private Sub fselect_Click(sender As Object, e As EventArgs) _
        Handles fselect.Click, Label1.Click, PictureBox1.Click

        isObjectSelectionActive = False

        Try
            If Globals.g_inventorApplication IsNot Nothing Then
                SetForegroundWindow(New IntPtr(Globals.g_inventorApplication.MainFrameHWND))
                SendKeys.SendWait("{ESC}")
            End If
        Catch
        End Try

        If m_selectedComponent Is Nothing OrElse m_centerPoint Is Nothing Then
            MessageBox.Show("먼저 TUBE 또는 PIPE 객체를 선택해야 합니다.", "안내")
            Me.BeginInvoke(New Action(AddressOf RunAnalysisSequence))
            Return
        End If

        Me.BeginInvoke(New Action(AddressOf RunSupportFaceSelection))
    End Sub

    Private Sub fselect_MouseEnter(sender As Object, e As EventArgs) _
        Handles fselect.MouseEnter, Label1.MouseEnter, PictureBox1.MouseEnter
        fselect.BackColor = cBtnHover
    End Sub

    Private Sub fselect_MouseLeave(sender As Object, e As EventArgs) _
        Handles fselect.MouseLeave, Label1.MouseLeave, PictureBox1.MouseLeave
        fselect.BackColor = cBtnNormal
    End Sub

    Private Sub fselect_MouseDown(sender As Object, e As MouseEventArgs) _
        Handles fselect.MouseDown, Label1.MouseDown, PictureBox1.MouseDown
        If e.Button = MouseButtons.Left Then
            fselect.BackColor = cBtnClick
        End If
    End Sub

    Private Sub fselect_MouseUp(sender As Object, e As MouseEventArgs) _
        Handles fselect.MouseUp, Label1.MouseUp, PictureBox1.MouseUp
        fselect.BackColor = cBtnHover
    End Sub

#End Region

    ' ========================= Part 2/2 =========================

#Region "I. Inventor 선택 루프 + 바닥면 선택/거리"

    Private Sub RunAnalysisSequence()
        If Globals.g_inventorApplication Is Nothing Then Return

        isObjectSelectionActive = True

        Dim selHL As HighlightSet = Nothing
        Try
            selHL = Globals.g_inventorApplication.ActiveDocument.CreateHighlightSet()
            selHL.Color = Globals.g_inventorApplication.TransientObjects.CreateColor(0, 191, 255)
        Catch
        End Try

        While isFormRunning AndAlso isObjectSelectionActive
            Dim oSelect As Object = Nothing
            Try
                SetForegroundWindow(New IntPtr(Globals.g_inventorApplication.MainFrameHWND))

                Dim pPick As Inventor.Point = Nothing
                Dim occPicked As ComponentOccurrence =
    PickOccurrenceWithPoint("Support 생성할 객체를 선택하세요", pPick)


                oSelect = occPicked
                m_lastPickPoint = pPick
            Catch
                Exit While
            End Try

            If oSelect Is Nothing OrElse Not TypeOf oSelect Is ComponentOccurrence Then Exit While

            Dim comp As ComponentOccurrence = CType(oSelect, ComponentOccurrence)
            Dim fileName As String = comp.Definition.Document.DisplayName.ToUpper()

            Try
                selHL?.Clear()
                selHL?.AddItem(comp)
            Catch
            End Try

            If fileName.Contains("TUBE") Then
                ttext.Text = "TUBE"
            ElseIf fileName.Contains("PIPE") Then
                ttext.Text = "PIPE"
            Else
                MessageBox.Show("TUBE 또는 PIPE만 선택 가능합니다.")
                Continue While
            End If

            Model.Text = "Line Support"

            SetSupportComponent(comp)

            Dim dia = GetMaxOuterDiameter(comp)
            stext.Text = GetSizeText(ttext.Text, dia)

            UpdateSupportLength()
            UpdateFilename()
        End While

        Try
            selHL?.Delete()
        Catch
        End Try
    End Sub

    Private Sub RunSupportFaceSelection()
        If m_centerPoint Is Nothing Then Return

        Try
            SetForegroundWindow(New IntPtr(Globals.g_inventorApplication.MainFrameHWND))
            Dim face = Globals.g_inventorApplication.CommandManager.Pick(
                SelectionFilterEnum.kPartFaceFilter,
                "지지할 바닥면 선택")

            If face Is Nothing Then Return

            Dim pickedObj As Object = face

            ' Face / FaceProxy 둘 다 허용
            Dim faceAsFace As Face = TryCast(pickedObj, Face)
            Dim faceAsProxy As FaceProxy = TryCast(pickedObj, FaceProxy)

            If faceAsFace Is Nothing AndAlso faceAsProxy Is Nothing Then Return

            ' m_supportFace는 Face 타입이라 Proxy면 GeometryProxy로 저장(후속 최소거리 계산용)
            If faceAsFace IsNot Nothing Then
                m_supportFace = faceAsFace
            Else
                m_supportFace = faceAsProxy   ' ★ FaceProxy도 Face처럼 취급 가능(변수 타입이 Face면 컴파일 에러 나면 아래 2)로 변경)
            End If

            ' [NEW] 바닥면 경사각(도) 계산 → af 텍스트박스 출력(소수점 둘째자리)
            Try
                If af IsNot Nothing Then
                    Dim slopeDeg As Double = GetFloorSlopeDeg(pickedObj)
                    af.Text = slopeDeg.ToString("0.00", Globalization.CultureInfo.InvariantCulture)

                    Try
                        Dim fp As FaceProxy = TryCast(pickedObj, FaceProxy)
                        Globals.g_inventorApplication.StatusBarText =
                            $"[SlopeDebug] isProxy={(fp IsNot Nothing)}  slope={af.Text}"
                    Catch
                    End Try
                End If
            Catch
            End Try

            ' 재질 읽기
            Dim matCode As String = GetMaterialFromPickedFace(pickedObj)
            If Not String.IsNullOrWhiteSpace(matCode) Then
                mtext.Text = matCode
            End If

            ApplyMaterialRules()   ' ★ 추가
            UpdateOptionButtons()
            UpdateFilename()

            Dim distCm =
    Globals.g_inventorApplication.MeasureTools.GetMinimumDistance(m_centerPoint, pickedObj)


            length.Text = (distCm * 10).ToString("0.00")
            UpdateSupportLength()
            UpdateFilename()

            Me.BeginInvoke(New Action(AddressOf RunPreviewPlacement))

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#End Region

#Region "J. SUPPORT 파일 탐색 규칙"

    Private Function NormalizeMaterialName(rawName As String) As String
        Dim s As String = If(rawName, "").Trim()
        If s = "" Then Return ""

        Dim idx As Integer = s.IndexOf("("c)
        If idx >= 0 Then s = s.Substring(0, idx)

        s = s.Trim().ToUpper()

        If s.Contains("STS") OrElse s.Contains("SUS") OrElse s.Contains("STAINLESS") Then Return "STS"
        If s.Contains("PVC") Then Return "PVC"

        Return s
    End Function

    Private Function GetMaterialFromPickedFace(faceObj As Object) As String
        Try
            If faceObj Is Nothing Then Return ""

            Dim matName As String = ""

            Dim fp As FaceProxy = TryCast(faceObj, FaceProxy)
            If fp IsNot Nothing Then
                Dim occ As ComponentOccurrence = fp.ContainingOccurrence
                If occ IsNot Nothing Then
                    Dim pdoc As PartDocument = TryCast(occ.Definition.Document, PartDocument)
                    If pdoc IsNot Nothing Then
                        matName = pdoc.ComponentDefinition.Material.Name
                    End If
                End If
            Else
                Dim f As Face = TryCast(faceObj, Face)
                If f IsNot Nothing Then
                    ' 이 루트는 불안정하니 Try 내에서만
                    Dim doc As Document = TryCast(f.Parent.Parent.Parent, Document)
                    Dim pdoc As PartDocument = TryCast(doc, PartDocument)
                    If pdoc IsNot Nothing Then
                        matName = pdoc.ComponentDefinition.Material.Name
                    End If
                End If
            End If

            Return NormalizeMaterialName(matName)
        Catch
            Return ""
        End Try
    End Function

    Private Const SUPPORT_BASE_PATH As String =
        "O:\System_BU_Vault\00. 구매품 Library\support\"

    Private Function GetSupportFolderPath() As String
        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

        Select Case mat
            Case "PVC"
                ' PVC: FootType에 따라 폴더 선택 (CF=Basic, GF=Grip)
                Dim ft As String = GetLegCode() ' CF / GF
                If ft = "GF" Then
                    Return IO.Path.Combine(SUPPORT_BASE_PATH, "PVC_SUPPORT", "Grip")
                Else
                    Return IO.Path.Combine(SUPPORT_BASE_PATH, "PVC_SUPPORT", "Basic")
                End If

            Case "STS"
                ' STS: 기존 유지(최상위 폴더)
                Return IO.Path.Combine(SUPPORT_BASE_PATH, "STS_SUPPORT")
        End Select
        Return ""
    End Function

    Private Function FindSupportFile() As String
        Dim folder As String = GetSupportFolderPath()
        If String.IsNullOrWhiteSpace(folder) Then Return ""
        If Not IO.Directory.Exists(folder) Then Return ""

        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

        If mat = "STS" Then
            Return FindStsFileByCoreKeys_IgnoreSizeLength(folder)
        End If

        If mat = "PVC" Then
            Return FindPvcFileByCoreKeys_SizeRequired_IgnoreLength(folder)
        End If
        Return ""

    End Function
    ' [NEW] PVC: Size/Length 무시, 폴더는 이미 Basic/Grip로 좁혀져 있음
    Private Function FindPvcFileByCoreKeys_SizeRequired_IgnoreLength(pvcFolder As String) As String
        If Not IO.Directory.Exists(pvcFolder) Then Return ""

        Dim sm As String = ConvertSingleMultiToCode()      ' S/M
        Dim typ As String = GetSadTypeCode() ' TU/PI
        Dim sizeCode As String = ConvertSizeToCode(stext.Text) ' ★ PVC는 사이즈 필수
        Dim leg As String = GetLegCode()                   ' CF/GF
        Dim bolt As String = "N"                           ' PVC 고정

        Dim files = IO.Directory.GetFiles(pvcFolder, "*.ipt", IO.SearchOption.TopDirectoryOnly)

        ' 1) PVC + S/M + TU/PI + SIZE + CF/GF + N
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If Not HasToken(n, "PVC") Then Continue For
            If Not HasToken(n, sm) Then Continue For
            If Not HasToken(n, typ) Then Continue For
            If Not HasToken(n, sizeCode) Then Continue For   ' ★ 추가
            If Not HasToken(n, leg) Then Continue For
            If Not HasToken(n, bolt) Then Continue For
            Return f
        Next

        ' 2) PVC + TU/PI + SIZE + CF/GF
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If Not HasToken(n, "PVC") Then Continue For
            If Not HasToken(n, typ) Then Continue For
            If Not HasToken(n, sizeCode) Then Continue For   ' ★ 추가
            If Not HasToken(n, leg) Then Continue For
            Return f
        Next

        ' 3) PVC + SIZE + CF/GF
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If Not HasToken(n, "PVC") Then Continue For
            If Not HasToken(n, sizeCode) Then Continue For   ' ★ 추가
            If Not HasToken(n, leg) Then Continue For
            Return f
        Next

        Return ""
    End Function

    Private Function PlaceSupportOccurrence(supportPath As String,
                                           Optional spawnPoint As Inventor.Point = Nothing) As ComponentOccurrence
        If Globals.g_inventorApplication Is Nothing Then Return Nothing

        Dim doc As Document = Globals.g_inventorApplication.ActiveDocument
        If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MessageBox.Show("어셈블리(.iam)에서만 배치할 수 있습니다.", "문서 오류")
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(supportPath) OrElse Not IO.File.Exists(supportPath) Then
            MessageBox.Show("배치할 파일 경로가 유효하지 않습니다." & vbCrLf & supportPath, "경로 오류")
            Return Nothing
        End If

        Dim asm As AssemblyDocument = CType(doc, AssemblyDocument)
        Dim asmDef As AssemblyComponentDefinition = asm.ComponentDefinition

        Try
            Dim m As Matrix = Globals.g_inventorApplication.TransientGeometry.CreateMatrix()
            m.SetToIdentity()

            If spawnPoint IsNot Nothing Then
                m.Cell(1, 4) = spawnPoint.X
                m.Cell(2, 4) = spawnPoint.Y
                m.Cell(3, 4) = spawnPoint.Z
            End If

            Dim occ As ComponentOccurrence = asmDef.Occurrences.Add(supportPath, m)
            Globals.g_inventorApplication.ActiveView.Update()
            Return occ
        Catch ex As Exception
            MessageBox.Show(ex.Message, "배치(Add) 오류")
            Return Nothing
        End Try
    End Function

    Private Sub UpdateSupportParameters(occ As ComponentOccurrence, sdMm As Double, splMm As Double, Optional setSD As Boolean = True, Optional doSave As Boolean = False)

        If occ Is Nothing Then Return

        Dim partDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
        If partDef Is Nothing Then Return

        Dim pdoc As PartDocument = TryCast(partDef.Document, PartDocument)
        If pdoc Is Nothing Then Return

        Dim sdExpr As String = sdMm.ToString("0.###", Globalization.CultureInfo.InvariantCulture) & " mm"
        Dim splExpr As String = splMm.ToString("0.###", Globalization.CultureInfo.InvariantCulture) & " mm"

        Dim okSD As Boolean = False
        Dim okSPL As Boolean = False

        Try
            Dim ups As UserParameters = partDef.Parameters.UserParameters
            If setSD Then
                Try
                    ups.Item("SD").Expression = sdExpr
                    okSD = True
                Catch
                End Try
            Else
                okSD = True ' PVC는 SD를 안 만지는 것이 정상 → OK 취급
            End If

            Try
                ups.Item("SPL").Expression = splExpr
                okSPL = True
            Catch
            End Try

        Catch
        End Try

        If (Not okSD) OrElse (Not okSPL) Then
            Try
                Dim mps As ModelParameters = partDef.Parameters.ModelParameters
                For Each p As Parameter In mps
                    If p Is Nothing OrElse p.Name Is Nothing Then Continue For
                    Dim nm As String = p.Name.Trim().ToUpper()

                    If setSD AndAlso (Not okSD) AndAlso nm = "SD" Then
                        Try : p.Expression = sdExpr : okSD = True : Catch : End Try
                    End If

                    If (Not okSPL) AndAlso nm = "SPL" Then
                        Try : p.Expression = splExpr : okSPL = True : Catch : End Try
                    End If

                    If okSD AndAlso okSPL Then Exit For
                Next
            Catch
            End Try
        End If

        Try
            Dim readSD As String = ""
            Dim readSPL As String = ""
            Try
                readSD = partDef.Parameters.Item("SD").Expression
            Catch
            End Try
            Try
                readSPL = partDef.Parameters.Item("SPL").Expression
            Catch
            End Try

            Globals.g_inventorApplication.StatusBarText =
                $"[SupportParam] SD={(If(okSD, "OK", "NG"))}({readSD})  SPL={(If(okSPL, "OK", "NG"))}({readSPL})  file={IO.Path.GetFileName(pdoc.FullFileName)}"
        Catch
        End Try

        Try : pdoc.Update2(True) : Catch : End Try

        If doSave Then
            Try : pdoc.Save2(True) : Catch : End Try
        End If
    End Sub


    Private Function HasToken(nameUpper As String, tokenUpper As String) As Boolean
        nameUpper = If(nameUpper, "").Trim().ToUpper()
        tokenUpper = If(tokenUpper, "").Trim().ToUpper()
        If nameUpper = "" OrElse tokenUpper = "" Then Return False

        Dim pattern As String =
            "(^|[^A-Z0-9])" &
            System.Text.RegularExpressions.Regex.Escape(tokenUpper) &
            "([^A-Z0-9]|$)"

        Return System.Text.RegularExpressions.Regex.IsMatch(nameUpper, pattern)
    End Function

    Private Function FindStsFileByCoreKeys_IgnoreSizeLength(stsFolder As String) As String
        If Not IO.Directory.Exists(stsFolder) Then Return ""

        Dim sm As String = ConvertSingleMultiToCode()
        Dim leg As String = GetLegCode()
        Dim bolt As String = ConvertBoltToCode()

        Dim files = IO.Directory.GetFiles(stsFolder, "*.ipt", IO.SearchOption.TopDirectoryOnly)

        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, sm) Then Continue For
            If Not HasToken(n, leg) Then Continue For
            If Not HasToken(n, bolt) Then Continue For
            Return f
        Next

        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, leg) Then Continue For
            If Not HasToken(n, bolt) Then Continue For
            Return f
        Next

        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, leg) Then Continue For
            Return f
        Next

        Return ""
    End Function

    Private Function ApplyTempAxisToAxisPreview(supportOcc As ComponentOccurrence,
                                            targetOcc As ComponentOccurrence,
                                            ByRef constraintStore As AssemblyConstraint,
                                            Optional preferAxisZName As Boolean = False,
                                            Optional spawnPoint As Inventor.Point = Nothing) As Boolean

        If Globals.g_inventorApplication Is Nothing Then Return False
        If supportOcc Is Nothing OrElse targetOcc Is Nothing Then Return False

        ' [FIX] tg 스코프 고정
        Dim app As Inventor.Application = Globals.g_inventorApplication
        Dim tg As Inventor.TransientGeometry = app.TransientGeometry

        Dim doc As Document = Globals.g_inventorApplication.ActiveDocument
        If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Return False

        Dim asm As AssemblyDocument = CType(doc, AssemblyDocument)
        Dim asmDef As AssemblyComponentDefinition = asm.ComponentDefinition

        Try
            If constraintStore IsNot Nothing Then constraintStore.Delete()
        Catch
        End Try
        constraintStore = Nothing

        Try
            supportOcc.Grounded = False
        Catch
        End Try

        Dim supportAxisProxy As WorkAxisProxy = Nothing
        If preferAxisZName Then
            supportAxisProxy = GetAxisZProxyPreferNamed(supportOcc, "AXIS_Z")
        Else
            supportAxisProxy = GetSupportOriginZAxisProxy(supportOcc)
        End If

        If supportAxisProxy Is Nothing Then
            MessageBox.Show("Support/SAD 파일에서 축을 찾을 수 없습니다.", "오류")
            Return False
        End If

        ' ★ 구속 전에, 서포트를 스폰점 근처로 먼저 붙여놓는다(날아감 감소)
        If spawnPoint IsNot Nothing Then
            NudgeOccurrenceAxisToPoint(supportOcc, supportAxisProxy, spawnPoint)
        End If

        Dim targetAxisProxy As WorkAxisProxy = GetTargetCylinderAxisProxyNearPick(targetOcc, m_lastPickPoint)
        ' [DEBUG] 타겟 축 프록시 생성 여부
        Try
            Dim cntCyl As Integer = 0
            Try
                cntCyl = GetCylinderFaceProxiesFromOccurrence(targetOcc).Count
            Catch
                cntCyl = -1
            End Try

            Globals.g_inventorApplication.StatusBarText =
        $"[AxisDebug] targetCylFaces={cntCyl}  targetAxisProxy={(If(targetAxisProxy IsNot Nothing, "OK", "NG"))}"
        Catch
        End Try


        If targetAxisProxy Is Nothing Then
            MessageBox.Show("타겟 파트에서 원통 축을 만들 수 없습니다.", "형상 찾기 실패")
            Return False
        End If

        Try
            ' ★ 구속 걸기 전에 방향/위치를 먼저 대충 맞춰서 솔버의 "다른 해" 점프를 막는다
            'PreAlignOccurrenceToTargetAxis(supportOcc, targetFaceProxy, supportAxisProxy, spawnPoint)

            ' =========================================================
            ' [CHANGED - TEST] (C) 축방향 잠금(WorkPlane 생성 + Flush) "끄기"
            ' - 아래 흐름만 유지:
            '   (A) 스폰점 WorkPoint 생성
            '   (B) Support WorkPointProxy ↔ 스폰 WorkPoint : Point-Point Mate
            '   (D) Support 축 ↔ Target 원통 축 : Axis Mate
            ' =========================================================

            Try
                ' --- 기존 잠금 구속/기하 제거 ---
                SafeRun(Sub()
                            If m_tempSpawnConstraint IsNot Nothing Then
                                m_tempSpawnConstraint.Delete()
                            End If
                        End Sub)
                m_tempSpawnConstraint = Nothing

                SafeRun(Sub()
                            If m_tempSpawnWorkPoint IsNot Nothing Then
                                m_tempSpawnWorkPoint.Delete()
                            End If
                        End Sub)
                m_tempSpawnWorkPoint = Nothing

                ' (C)에서 쓰던 것들도 혹시 남아있으면 삭제만 해둔다
                SafeRun(Sub()
                            If m_tempAxialLockConstraint IsNot Nothing Then
                                m_tempAxialLockConstraint.Delete()
                            End If
                        End Sub)
                m_tempAxialLockConstraint = Nothing

                SafeRun(Sub()
                            If m_tempAxialLockPlane IsNot Nothing Then
                                m_tempAxialLockPlane.Delete()
                            End If
                        End Sub)
                m_tempAxialLockPlane = Nothing

                If spawnPoint IsNot Nothing Then

                    ' (A) WorkPoint 고정 생성
                    m_tempSpawnWorkPoint = asmDef.WorkPoints.AddFixed(spawnPoint)

                    ' (B) ★점-점 Mate로 "위치"를 고정 (Support WorkPoint ↔ 스폰 WorkPoint)
                    Dim supportOriginWpProxy As WorkPointProxy = GetSupportWorkPointProxyOnAxis(supportOcc, supportAxisProxy)


                    If supportOriginWpProxy IsNot Nothing Then
                        m_tempSpawnConstraint =
                        asmDef.Constraints.AddMateConstraint(
                            supportOriginWpProxy,
                            m_tempSpawnWorkPoint,
                            0.0,
                            InferredTypeEnum.kInferredPoint,
                            InferredTypeEnum.kInferredPoint
                        )
                    Else
                        Return False
                    End If

                    ' (C) 축방향 잠금(평면+Flush) — 여기서 "완전히 생략"
                End If
            Catch
            End Try

            ' (D) 축-축 Mate (마지막)
            constraintStore = asmDef.Constraints.AddMateConstraint(
    supportAxisProxy,
    targetAxisProxy,
    0.0,
    InferredTypeEnum.kInferredLine,
    InferredTypeEnum.kInferredLine)


            Globals.g_inventorApplication.ActiveView.Update()
            Return (constraintStore IsNot Nothing)

        Catch ex As Exception
            MessageBox.Show("구속 실패: " & ex.Message & vbCrLf & "원통면 중심축 인식 실패", "구속 오류")
            Return False
        End Try

    End Function

    Private Function GetSupportOriginZAxisProxy(occ As ComponentOccurrence) As WorkAxisProxy
        Try
            If occ Is Nothing Then Return Nothing

            Dim zAxis As WorkAxis = occ.Definition.WorkAxes.Item(3)

            Dim proxyObj As Object = Nothing
            occ.CreateGeometryProxy(zAxis, proxyObj)
            Return TryCast(proxyObj, WorkAxisProxy)
        Catch
            Return Nothing
        End Try
    End Function
    ' =========================
    ' [NEW] Occurrence에서 원통 FaceProxy를 안전하게 수집(조립품 컨텍스트 우선)
    ' - 1순위: occ.SurfaceBodies (SurfaceBodyProxy)에서 FaceProxy를 직접 얻음
    ' - 2순위: occ.Definition.SurfaceBodies → CreateGeometryProxy로 FaceProxy 생성
    ' =========================
    Private Function GetCylinderFaceProxiesFromOccurrence(occ As ComponentOccurrence) As List(Of FaceProxy)
        Dim result As New List(Of FaceProxy)

        Try
            If occ Is Nothing Then Return result

            ' 1) Assembly 컨텍스트(Proxy) 우선: occ.SurfaceBodies
            Try
                Dim sbs As SurfaceBodies = occ.SurfaceBodies
                If sbs IsNot Nothing AndAlso sbs.Count > 0 Then
                    For Each sb As SurfaceBody In sbs
                        If sb Is Nothing Then Continue For
                        For Each f As Face In sb.Faces
                            If f Is Nothing Then Continue For
                            If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                                Dim fp As FaceProxy = TryCast(f, FaceProxy)
                                If fp IsNot Nothing Then result.Add(fp)
                            End If
                        Next
                    Next
                End If
            Catch
                ' 무시하고 2)로
            End Try

            ' 2) 그래도 없으면: Definition(로컬) → CreateGeometryProxy
            If result.Count = 0 Then
                Dim partDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
                If partDef IsNot Nothing Then
                    For Each body As SurfaceBody In partDef.SurfaceBodies
                        If body Is Nothing Then Continue For
                        For Each f As Face In body.Faces
                            If f Is Nothing Then Continue For
                            If f.SurfaceType <> SurfaceTypeEnum.kCylinderSurface Then Continue For

                            Dim prxObj As Object = Nothing
                            occ.CreateGeometryProxy(f, prxObj)
                            Dim fp As FaceProxy = TryCast(prxObj, FaceProxy)
                            If fp IsNot Nothing Then result.Add(fp)
                        Next
                    Next
                End If
            End If

        Catch
        End Try

        Return result
    End Function


    ' =========================
    ' [CHANGED] 클릭점(pickPt)에 가장 가까운 "원통 FaceProxy"를 찾는다
    ' - 변경 포인트: occ.Definition.SurfaceBodies 직접 루프 → (위 NEW) FaceProxy 수집 함수 사용
    ' =========================
    Private Function GetTargetCylinderFaceProxyNearPick(occ As ComponentOccurrence,
                                                    pickPt As Inventor.Point) As FaceProxy
        Try
            If Globals.g_inventorApplication Is Nothing Then Return Nothing
            If occ Is Nothing Then Return Nothing
            If pickPt Is Nothing Then Return Nothing

            Dim app = Globals.g_inventorApplication

            Dim candidates As List(Of FaceProxy) = GetCylinderFaceProxiesFromOccurrence(occ)
            If candidates Is Nothing OrElse candidates.Count = 0 Then Return Nothing

            Dim best As FaceProxy = Nothing
            Dim bestDist As Double = Double.MaxValue

            For Each fp As FaceProxy In candidates
                If fp Is Nothing Then Continue For

                Dim d As Double = Double.MaxValue
                Try
                    d = app.MeasureTools.GetMinimumDistance(pickPt, fp)
                Catch
                    Continue For
                End Try

                If d < bestDist Then
                    bestDist = d
                    best = fp
                End If
            Next

            Return best

        Catch
            Return Nothing
        End Try
    End Function


    ' =========================
    ' [CHANGED] 타겟 Occurrence에서 "대표 원통 FaceProxy" 하나를 반환(반지름 최대)
    ' - 변경 포인트: Definition 직접 루프 → FaceProxy 수집 후, Geometry를 Cylinder로 캐스팅
    ' =========================
    Private Function GetTargetCylinderFaceProxy(occ As ComponentOccurrence) As FaceProxy
        Try
            If occ Is Nothing Then Return Nothing

            Dim candidates As List(Of FaceProxy) = GetCylinderFaceProxiesFromOccurrence(occ)
            If candidates Is Nothing OrElse candidates.Count = 0 Then Return Nothing

            Dim best As FaceProxy = Nothing
            Dim bestR As Double = -1.0

            For Each fp As FaceProxy In candidates
                If fp Is Nothing Then Continue For

                Dim cyl As Cylinder = Nothing
                Try
                    cyl = TryCast(fp.Geometry, Cylinder)
                Catch
                    cyl = Nothing
                End Try
                If cyl Is Nothing Then Continue For

                If cyl.Radius > bestR Then
                    bestR = cyl.Radius
                    best = fp
                End If
            Next

            Return best

        Catch
            Return Nothing
        End Try
    End Function

    Private Function FindStsSadFile(stsFolder As String) As String
        If Not IO.Directory.Exists(stsFolder) Then Return ""

        Dim sm As String = ConvertSingleMultiToCode()
        Dim sizeCode As String = ConvertSizeToCode(stext.Text)

        Dim files = IO.Directory.GetFiles(stsFolder, "*.ipt", IO.SearchOption.TopDirectoryOnly)

        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, sm) Then Continue For
            If Not HasToken(n, "SAD") Then Continue For
            If Not HasToken(n, sizeCode) Then Continue For

            Return f
        Next

        ' fallback: 사이즈만 맞는 SAD
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, "SAD") Then Continue For
            If Not HasToken(n, sizeCode) Then Continue For

            Return f
        Next

        Return ""
    End Function


    Private Function GetAxisZProxyPreferNamed(occ As ComponentOccurrence, Optional axisName As String = "AXIS_Z") As WorkAxisProxy
        Try
            If occ Is Nothing Then Return Nothing

            Dim wa As WorkAxis = Nothing

            Try
                For Each a As WorkAxis In occ.Definition.WorkAxes
                    If a IsNot Nothing AndAlso a.Name IsNot Nothing Then
                        If a.Name.Trim().ToUpper() = axisName.Trim().ToUpper() Then
                            wa = a
                            Exit For
                        End If
                    End If
                Next
            Catch
            End Try

            If wa Is Nothing Then
                wa = occ.Definition.WorkAxes.Item(3)
            End If

            Dim proxyObj As Object = Nothing
            occ.CreateGeometryProxy(wa, proxyObj)
            Return TryCast(proxyObj, WorkAxisProxy)
        Catch
            Return Nothing
        End Try
    End Function

#End Region

#Region "K. 파일명 생성 규칙(UpdateFilename 실시간)"

    ' [NEW] 재질 규칙 반영: PVC는 Bolt=N 고정/잠금
    Private Sub ApplyMaterialRules()
        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

        If mat = "PVC" Then
            If btype IsNot Nothing Then
                btype.Checked = False
                btype.Enabled = False
                btype.BackColor = Draw.Color.FromArgb(60, 60, 60)
            End If

            ' PVC 기본 CF
            If String.IsNullOrWhiteSpace(m_legCodePVC) Then m_legCodePVC = "CF"
        Else
            If btype IsNot Nothing Then
                btype.Enabled = True
                btype.BackColor = cBtnNormal
            End If

            ' STS 기본 SF
            If String.IsNullOrWhiteSpace(m_legCode) Then m_legCode = "SF"
        End If

    End Sub


    Private Function ConvertLengthToFileToken(lenText As String) As String
        Dim raw As String = If(lenText, "").Trim()
        If raw = "" Then Return "L0"

        Dim v As Double = 0

        If Not Double.TryParse(raw,
                               Globalization.NumberStyles.Any,
                               Globalization.CultureInfo.InvariantCulture,
                               v) Then
            If Not Double.TryParse(raw, v) Then
                Return "L0"
            End If
        End If

        Dim s As String = v.ToString("0.###", Globalization.CultureInfo.InvariantCulture)
        s = s.Replace(".", "P")
        Return "L" & s
    End Function

    Private Function ConvertSizeToCode(sizeText As String) As String
        sizeText = If(sizeText, "").Trim()
        If sizeText = "" Then Return "000"

        Select Case sizeText.ToUpper()
            Case "GEN" : Return "000"
            Case "6A" : Return "06A"
            Case "10A" : Return "10A"
            Case "15A" : Return "15A"
            Case "20A" : Return "20A"
            Case "25A" : Return "25A"
            Case "40A" : Return "40A"
            Case "50A" : Return "50A"
        End Select

        Select Case sizeText
            Case "1/8""" : Return "04B"
            Case "1/4""" : Return "06B"
            Case "3/8""" : Return "10B"
            Case "1/2""" : Return "15B"
            Case "3/4""" : Return "20B"
            Case "1""" : Return "25B"
            Case "11/2""" : Return "40B"
            Case "2""" : Return "50B"
        End Select

        Return "000"
    End Function

    Private Function ConvertTypeToCode(typeText As String) As String
        Select Case If(typeText, "").Trim().ToUpper()
            Case "TUBE" : Return "TU"
            Case "PIPE" : Return "PI"
        End Select
        Return ""
    End Function

    Private Function ConvertModelToCode(modelText As String) As String
        If If(modelText, "").Trim().ToUpper() = "LINE SUPPORT" Then Return "LSP"
        Return ""
    End Function

    Private Function ConvertSingleMultiToCode() As String
        If mbt IsNot Nothing AndAlso mbt.Checked Then Return "M"
        Return "S"
    End Function

    Private Function ConvertBoltToCode() As String
        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()
        If mat = "PVC" Then Return "N" ' PVC 고정

        If btype IsNot Nothing AndAlso btype.Checked Then Return "Y"
        Return "N"
    End Function


    Private Function GetLegCode() As String
        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

        If mat = "STS" Then
            If String.IsNullOrWhiteSpace(m_legCode) Then m_legCode = "SF"
            Return m_legCode
        End If

        If mat = "PVC" Then
            If String.IsNullOrWhiteSpace(m_legCodePVC) Then m_legCodePVC = "CF"
            Return m_legCodePVC
        End If

        Return "GEN"
    End Function


    Private Sub UpdateFilename()
        If filename Is Nothing Then Exit Sub

        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()
        Dim sm As String = ConvertSingleMultiToCode()
        Dim typ As String = ConvertTypeToCode(ttext.Text)
        Dim mdl As String = ConvertModelToCode(Model.Text)
        Dim sizeCode As String = ConvertSizeToCode(stext.Text)
        Dim leg As String = GetLegCode()
        Dim bolt As String = ConvertBoltToCode()
        Dim lenPart As String = ConvertLengthToFileToken(splength.Text)

        If mat = "" OrElse typ = "" OrElse mdl = "" OrElse sizeCode = "" Then
            filename.Text = ""
            Exit Sub
        End If

        filename.Text = $"{mat}-{sm}-{typ}-{mdl}-{sizeCode}-{leg}-{bolt}-{lenPart}"
    End Sub

#End Region

#Region "L. splength 계산 + Distance MouseWheel"

    Private Sub Distance_MouseWheel(sender As Object, e As MouseEventArgs)
        Dim v As Double = 0
        Double.TryParse(Distance.Text, v)

        If e.Delta > 0 Then
            v += 1.0
        Else
            v -= 1.0
        End If

        Distance.Text = v.ToString("0.00")
    End Sub

    Private Function GetSizeOffset(sizeText As String) As Double
        Select Case sizeText
            Case "1""" : Return 12.4
            Case "3/4""" : Return 9.2
            Case "1/2""" : Return 6.1
            Case "3/8""" : Return 4.5
            Case "1/4""" : Return 2.9
            Case "20A" : Return 13.0
            Case "25A" : Return 16.4
        End Select
        Return 0
    End Function

    Private Sub UpdateSupportLength()
        If splength Is Nothing Then Exit Sub

        Dim baseLen As Double
        If Not Double.TryParse(length.Text, baseLen) Then
            splength.Text = ""
            Exit Sub
        End If

        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

        If mat = "PVC" Then
            ' PVC: 보정 금지 → splength = length
            splength.Text = Math.Max(0, baseLen) _
        .ToString("0.00", Globalization.CultureInfo.InvariantCulture)
        Else
            ' STS: 기존 보정 유지
            splength.Text = Math.Max(0, baseLen - GetSizeOffset(stext.Text)) _
        .ToString("0.00", Globalization.CultureInfo.InvariantCulture)
        End If

        UpdateFilename()
    End Sub

#End Region

#Region "M. Geometry Helpers(외경/사이즈/오차) + PickOccurrenceWithPoint"
    ' [NEW] 바닥면(평면) 경사각(도) 계산: 전역 Z축 대비 0~90°
    Private Function GetFloorSlopeDeg(pickedFaceObj As Object) As Double
        Try
            Dim app = Globals.g_inventorApplication
            If app Is Nothing Then Return 0

            Dim fp As FaceProxy = TryCast(pickedFaceObj, FaceProxy)
            Dim f As Face = TryCast(pickedFaceObj, Face)

            Dim geom As Object = Nothing
            Dim occ As ComponentOccurrence = Nothing

            If fp IsNot Nothing Then
                geom = fp.Geometry
                occ = fp.ContainingOccurrence
            ElseIf f IsNot Nothing Then
                geom = f.Geometry
            Else
                Return 0
            End If

            Dim pl As Plane = TryCast(geom, Plane)
            If pl Is Nothing Then Return 0

            Dim n As Vector = pl.Normal
            If n Is Nothing Then Return 0

            ' ★ 핵심: 어셈블리에서 선택한 FaceProxy면, 노멀을 Occurrence 변환으로 회전시켜 전역 기준으로 맞춘다
            If occ IsNot Nothing Then
                Try
                    Dim nnv As Vector = app.TransientGeometry.CreateVector(n.X, n.Y, n.Z)
                    nnv.TransformBy(occ.Transformation)   ' Vector라 translation은 영향 거의 없고, 회전만 반영됨
                    n = nnv
                Catch
                End Try
            End If

            Dim z As Vector = app.TransientGeometry.CreateVector(0, 0, 1)

            Dim dot As Double = Math.Abs(n.DotProduct(z))
            Dim nLen As Double = n.Length
            If nLen <= 0 Then Return 0

            Dim cosv As Double = dot / nLen   ' z 길이는 1이라 생략 가능
            If cosv > 1 Then cosv = 1
            If cosv < -1 Then cosv = -1

            Dim deg As Double = Math.Acos(cosv) * 180.0 / Math.PI
            Return deg
        Catch
            Return 0
        End Try
    End Function





    Private Function GetMaxOuterDiameter(comp As ComponentOccurrence) As Double
        Try
            Dim maxDia As Double = 0

            For Each body As SurfaceBody In comp.Definition.SurfaceBodies
                For Each f As Face In body.Faces
                    If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        Dim cyl As Cylinder = CType(f.Geometry, Cylinder)
                        Dim dia As Double = cyl.Radius * 20.0 ' cm 기준 → mm
                        If dia > maxDia Then maxDia = dia
                    End If
                Next
            Next

            Return maxDia
        Catch
            Return 0
        End Try
    End Function

    Private Function GetSizeText(typeText As String, dia As Double) As String
        ' ★ TYPE이 PIPE로 잡혀도, 실제 외경이 튜브 규격이면 인치로 먼저 매핑
        If IsInRange(dia, 25.4, 1.0) Then Return "1"""
        If IsInRange(dia, 19.05, 1.0) Then Return "3/4"""
        If IsInRange(dia, 12.7, 1.0) Then Return "1/2"""
        If IsInRange(dia, 9.53, 1.0) Then Return "3/8"""
        If IsInRange(dia, 6.35, 1.0) Then Return "1/4"""
        If IsInRange(dia, 4.76, 1.0) Then Return "3/16"""
        If IsInRange(dia, 3.18, 1.0) Then Return "1/8"""

        ' 그 다음 PIPE(A) 규격 매핑
        If IsInRange(dia, 60.33, 1.0) Then Return "50A"
        If IsInRange(dia, 48.26, 1.0) Then Return "40A"
        If IsInRange(dia, 33.4, 1.0) Then Return "25A"
        If IsInRange(dia, 26.67, 1.0) Then Return "20A"
        If IsInRange(dia, 21.34, 1.0) Then Return "15A"

        ' 매칭 실패면 mm로
        Return Math.Round(dia, 2).ToString("0.##", Globalization.CultureInfo.InvariantCulture) & "mm"
    End Function


    Private Function IsInRange(value As Double, target As Double, tol As Double) As Boolean
        Return Math.Abs(value - target) <= tol
    End Function
    ' [CHANGED] Occurrence + 클릭 3D 포인트(ModelPosition)까지 얻는 Pick
    ' - CommandManager.Pick은 Point를 주지 않아서 InteractionEvents.SelectEvents로 받는다
    Private Function PickOccurrenceWithPoint(prompt As String, ByRef pickedPt As Inventor.Point) As ComponentOccurrence
        pickedPt = Nothing
        If Globals.g_inventorApplication Is Nothing Then Return Nothing

        Dim app = Globals.g_inventorApplication

        Dim pickedOcc As ComponentOccurrence = Nothing
        Dim gotOne As Boolean = False

        ' [NEW] ByRef(pickedPt)를 람다에서 직접 건드리지 않기 위한 임시 변수
        Dim tempPickedPt As Inventor.Point = Nothing

        Dim ie As InteractionEvents = Nothing
        Dim se As SelectEvents = Nothing

        Try
            ie = app.CommandManager.CreateInteractionEvents()
            se = ie.SelectEvents

            ' 선택 필터: 조립품의 Leaf Occurrence
            se.ClearSelectionFilter()
            se.AddSelectionFilter(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter)
            se.SingleSelectEnabled = True

            ' 프롬프트 표시
            ie.StatusBarText = prompt

            AddHandler se.OnSelect,
        Sub(JustSelectedEntities As ObjectsEnumerator,
            SelectionDevice As SelectionDeviceEnum,
            ModelPosition As Point,
            ViewPosition As Point2d,
            View As Inventor.View)

            Try
                If JustSelectedEntities Is Nothing OrElse JustSelectedEntities.Count < 1 Then Return

                Dim o As Object = JustSelectedEntities.Item(1)
                Dim occ As ComponentOccurrence = TryCast(o, ComponentOccurrence)
                If occ Is Nothing Then Return

                pickedOcc = occ
                tempPickedPt = ModelPosition  ' ★ 여기만 로컬 변수로 저장
                gotOne = True
            Catch
            End Try
        End Sub

            ie.Start()

            ' 유저가 하나 고를 때까지 대기
            While Not gotOne
                System.Windows.Forms.Application.DoEvents()
                If Not isFormRunning Then Exit While
            End While

        Catch
        Finally
            Try
                If ie IsNot Nothing Then ie.Stop()
            Catch
            End Try
            Try
                ie = Nothing
                se = Nothing
            Catch
            End Try
        End Try

        ' [NEW] 루프 종료 후에만 ByRef에 대입 (BC36638 해결)
        pickedPt = tempPickedPt


        Return pickedOcc
    End Function
    ' [NEW] Occurrence에서 "대표 원통"의 중심축(Line)을 어셈블리 좌표로 만든다
    Private Function GetOccurrenceMainCylinderAxisLine(occ As ComponentOccurrence) As Inventor.Line
        Try
            If Globals.g_inventorApplication Is Nothing Then Return Nothing
            If occ Is Nothing Then Return Nothing

            Dim app = Globals.g_inventorApplication
            Dim tg = app.TransientGeometry

            Dim bestFace As Face = Nothing
            Dim maxRadius As Double = -1.0

            For Each body As SurfaceBody In occ.Definition.SurfaceBodies
                For Each f As Face In body.Faces
                    If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then
                        Dim cyl As Cylinder = CType(f.Geometry, Cylinder)
                        If cyl.Radius > maxRadius Then
                            maxRadius = cyl.Radius
                            bestFace = f
                        End If
                    End If
                Next
            Next

            If bestFace Is Nothing Then Return Nothing

            Dim bestCyl As Cylinder = CType(bestFace.Geometry, Cylinder)

            ' 로컬(파트) 축: BasePoint + AxisVector
            Dim p0 As Inventor.Point = tg.CreatePoint(bestCyl.BasePoint.X, bestCyl.BasePoint.Y, bestCyl.BasePoint.Z)
            Dim v0 As Inventor.Vector = tg.CreateVector(bestCyl.AxisVector.X, bestCyl.AxisVector.Y, bestCyl.AxisVector.Z)

            ' 어셈블리 좌표로 변환(회전+이동)
            p0.TransformBy(occ.Transformation)
            v0.TransformBy(occ.Transformation)

            Return tg.CreateLine(p0, v0)

        Catch
            Return Nothing
        End Try
    End Function

    ' [NEW] 클릭점(P)을 축 라인에 직교투영해서 P′를 만든다
    Private Function ProjectPointToLine(pt As Inventor.Point, axisLine As Inventor.Line) As Inventor.Point
        Try
            If Globals.g_inventorApplication Is Nothing Then Return Nothing
            If pt Is Nothing OrElse axisLine Is Nothing Then Return Nothing

            Dim app = Globals.g_inventorApplication
            Dim tg = app.TransientGeometry

            Dim p0 As Inventor.Point = axisLine.RootPoint
            Dim uv As UnitVector = TryCast(axisLine.Direction, UnitVector)
            If uv Is Nothing Then Return tg.CreatePoint(p0.X, p0.Y, p0.Z)

            Dim w As Inventor.Vector = tg.CreateVector(pt.X - p0.X, pt.Y - p0.Y, pt.Z - p0.Z)

            ' UnitVector는 길이가 1이라 /vv 필요 없음
            Dim t As Double = w.DotProduct(uv.AsVector())


            Return tg.CreatePoint(
     p0.X + uv.X * t,
     p0.Y + uv.Y * t,
     p0.Z + uv.Z * t
 )

        Catch
            Return Nothing
        End Try
    End Function

    ' ============================================================
    ' [CHANGED] 현재 선택된 타겟 + 마지막 클릭점 기준으로 "스폰점"을 만든다
    ' - 기존: 대표(최대) 원통 축 기준 → 클릭한 구간이 달라지면 스폰이 엇나감
    ' - 변경: "클릭점에 가장 가까운 원통 FaceProxy"의 축 기준으로 투영
    ' ============================================================
    Private Function GetSpawnPointOnTargetAxis() As Inventor.Point
        Try
            Dim app = Globals.g_inventorApplication
            If app Is Nothing Then Return m_centerPoint
            If m_selectedComponent Is Nothing Then Return m_centerPoint

            ' 1) 기준점: 클릭점 우선(없으면 center)
            Dim basePoint As Inventor.Point = If(m_lastPickPoint, m_centerPoint)
            If basePoint Is Nothing Then Return m_centerPoint

            ' 2) ★ 클릭점에 가장 가까운 원통 FaceProxy 선택
            Dim fp As FaceProxy = GetTargetCylinderFaceProxyNearPick(m_selectedComponent, basePoint)
            If fp Is Nothing Then
                app.StatusBarText = "[SpawnDebug] 클릭 근처 원통 FaceProxy 없음 → 중심점 사용"
                Return m_centerPoint
            End If

            ' 3) FaceProxy.Geometry는 Assembly 컨텍스트의 Cylinder여야 함
            Dim cyl As Cylinder = Nothing
            Try
                cyl = TryCast(fp.Geometry, Cylinder)
            Catch
                cyl = Nothing
            End Try

            If cyl Is Nothing Then
                app.StatusBarText = "[SpawnDebug] Cylinder 캐스팅 실패 → 중심점 사용"
                Return m_centerPoint
            End If

            ' 4) ★ 이 축(Line)은 이미 Assembly 좌표계 기준
            Dim axisLine As Inventor.Line = cyl.Axis
            If axisLine Is Nothing Then
                app.StatusBarText = "[SpawnDebug] 축(Line) 없음 → 중심점 사용"
                Return m_centerPoint
            End If

            ' 5) 클릭점을 축에 직교투영 → 최종 스폰점
            Dim proj As Inventor.Point = ProjectPointToLine(basePoint, axisLine)
            If proj IsNot Nothing Then
                app.StatusBarText =
                $"[SpawnDebug] OK (Pick->Axis) : ({proj.X:F3}, {proj.Y:F3}, {proj.Z:F3})"
                Return proj
            End If

            app.StatusBarText = "[SpawnDebug] 투영 실패 → 중심점 사용"
            Return m_centerPoint

        Catch ex As Exception
            Try
                Globals.g_inventorApplication.StatusBarText = "[SpawnDebug] 오류: " & ex.Message
            Catch
            End Try
            Return m_centerPoint
        End Try
    End Function



    ' [NEW] (Replace 성공 후) 이전 미리보기 임시파일 정리
    Private Sub CleanupPrevPreviewTempFile(prevPath As String, newPath As String)
        Try
            If String.IsNullOrWhiteSpace(prevPath) Then Return
            If String.IsNullOrWhiteSpace(newPath) Then Return
            If String.Equals(prevPath.Trim(), newPath.Trim(), StringComparison.OrdinalIgnoreCase) Then Return

            ' ★ prevPath가 "이번 세션 임시 생성" 목록에 있을 때만 삭제 (재사용 파일 보호)
            If Not m_tempCreatedFiles.Contains(prevPath) Then Return
            ' Inventor가 잡고 있을 수 있으니 닫고 지움
            TryCloseDocumentIfOpen(prevPath)

            Try : IO.File.SetAttributes(prevPath, IO.FileAttributes.Normal) : Catch : End Try

            For i As Integer = 1 To 5
                Try
                    IO.File.Delete(prevPath)
                    Exit For
                Catch
                    Try : System.Threading.Thread.Sleep(150) : Catch : End Try
                End Try
            Next

            ' 목록에서도 제거(중복 누적 방지)
            Try : m_tempCreatedFiles.Remove(prevPath) : Catch : End Try

        Catch
        End Try
    End Sub
    ' [기존] GetSupportOriginWorkPointProxy(...)
    ' → [변경] 스폰점에 가장 가까운 WorkPoint를 찾아 Proxy 반환

    Private Function GetSupportOriginWorkPointProxy(occ As ComponentOccurrence,
                                               Optional nearPt As Inventor.Point = Nothing) As WorkPointProxy
        Try
            If occ Is Nothing Then Return Nothing

            Dim app = Globals.g_inventorApplication
            If app Is Nothing Then Return Nothing

            Dim tg = app.TransientGeometry

            Dim partDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
            If partDef Is Nothing Then Return Nothing

            Dim bestWp As WorkPoint = Nothing
            Dim bestD As Double = Double.MaxValue

            ' nearPt가 없으면 기존처럼 1번을 우선 시도
            If nearPt Is Nothing Then
                Try
                    Dim wp1 As WorkPoint = partDef.WorkPoints.Item(1)
                    If wp1 IsNot Nothing Then
                        Dim prxObj1 As Object = Nothing
                        occ.CreateGeometryProxy(wp1, prxObj1)
                        Return TryCast(prxObj1, WorkPointProxy)
                    End If
                Catch
                End Try
                Return Nothing
            End If

            ' nearPt 기준으로 "가장 가까운 WorkPoint" 선택
            For Each wp As WorkPoint In partDef.WorkPoints
                If wp Is Nothing Then Continue For

                Dim p As Inventor.Point = Nothing
                Try
                    p = TryCast(wp.Point, Inventor.Point)
                Catch
                    p = Nothing
                End Try
                If p Is Nothing Then Continue For

                ' 파트 좌표 → 어셈블리 좌표
                Dim pg As Inventor.Point = tg.CreatePoint(p.X, p.Y, p.Z)
                pg.TransformBy(occ.Transformation)

                Dim dx = pg.X - nearPt.X
                Dim dy = pg.Y - nearPt.Y
                Dim dz = pg.Z - nearPt.Z
                Dim d2 = dx * dx + dy * dy + dz * dz

                If d2 < bestD Then
                    bestD = d2
                    bestWp = wp
                End If
            Next

            If bestWp Is Nothing Then Return Nothing

            Dim prxObj As Object = Nothing
            occ.CreateGeometryProxy(bestWp, prxObj)
            Return TryCast(prxObj, WorkPointProxy)

        Catch
            Return Nothing
        End Try
    End Function

    Private Sub PreAlignOccurrenceToTargetAxis(
    occ As ComponentOccurrence,
    targetCylFaceProxy As FaceProxy,
    occAxisProxy As WorkAxisProxy,
    spawnPoint As Inventor.Point
)
        Try
            If Globals.g_inventorApplication Is Nothing Then Exit Sub
            If occ Is Nothing OrElse targetCylFaceProxy Is Nothing OrElse occAxisProxy Is Nothing Then Exit Sub
            If spawnPoint Is Nothing Then Exit Sub

            Dim app As Inventor.Application = Globals.g_inventorApplication
            Dim tg As Inventor.TransientGeometry = app.TransientGeometry

            ' 1) target 축 방향(UnitVector) 구하기
            Dim cyl As Cylinder = TryCast(targetCylFaceProxy.Geometry, Cylinder)
            If cyl Is Nothing Then Exit Sub

            Dim nVec As Vector = cyl.AxisVector
            If nVec Is Nothing OrElse nVec.Length <= 0 Then Exit Sub
            Dim nUnit As UnitVector = tg.CreateUnitVector(nVec.X, nVec.Y, nVec.Z)

            ' 2) Support(occ) 축 방향(UnitVector) 구하기
            '    WorkAxisProxy.Geometry는 Line 일 때가 많다.
            Dim ln As Inventor.Line = TryCast(occAxisProxy.Geometry, Inventor.Line)
            If ln Is Nothing Then
                ' 혹시 Line이 아니면 그냥 탈출(프리정렬 생략)
                Exit Sub
            End If

            Dim sVec As Vector = ln.Direction
            If sVec Is Nothing OrElse sVec.Length <= 0 Then Exit Sub
            Dim sUnit As UnitVector = tg.CreateUnitVector(sVec.X, sVec.Y, sVec.Z)

            ' 3) 이미 거의 평행이면 회전 불필요
            Dim dp As Double = sUnit.DotProduct(nUnit)
            If dp > 1 Then dp = 1
            If dp < -1 Then dp = -1
            If Math.Abs(dp) > 0.9999 Then
                ' 방향만 거의 일치 → 위치만 스폰점으로 이동
                Dim mT As Matrix = occ.Transformation
                mT.Cell(1, 4) = spawnPoint.X
                mT.Cell(2, 4) = spawnPoint.Y
                mT.Cell(3, 4) = spawnPoint.Z
                occ.Transformation = mT
                Exit Sub
            End If

            ' 4) 회전축 = sUnit x nUnit, 회전각 = acos(dot)
            Dim rotAxisVec As Vector = sUnit.CrossProduct(nUnit)
            If rotAxisVec Is Nothing OrElse rotAxisVec.Length <= 0 Then Exit Sub
            Dim rotAxis As UnitVector = tg.CreateUnitVector(rotAxisVec.X, rotAxisVec.Y, rotAxisVec.Z)

            Dim ang As Double = Math.Acos(dp)

            ' 5) 현재 occ 변환에 회전 적용(스폰점 기준으로 회전)
            Dim mRot As Matrix = tg.CreateMatrix()
            mRot.SetToRotation(ang, rotAxis, spawnPoint)

            Dim m As Matrix = occ.Transformation
            m.TransformBy(mRot)

            ' 6) 마지막으로 위치를 스폰점으로 "확정"
            m.Cell(1, 4) = spawnPoint.X
            m.Cell(2, 4) = spawnPoint.Y
            m.Cell(3, 4) = spawnPoint.Z

            occ.Transformation = m

        Catch
        End Try
    End Sub


    ' [NEW] Occurrence의 축(Line)이 spawnPt를 지나가도록 "미리" 살짝 이동시킨다
    ' - 구속 걸 때 솔버가 멀리 점프하는 걸 크게 줄여줌
    Private Sub NudgeOccurrenceAxisToPoint(occ As ComponentOccurrence,
                                      axisProxy As WorkAxisProxy,
                                      spawnPt As Inventor.Point)
        Try
            If Globals.g_inventorApplication Is Nothing Then Return
            If occ Is Nothing OrElse axisProxy Is Nothing OrElse spawnPt Is Nothing Then Return

            Dim app = Globals.g_inventorApplication
            Dim tg = app.TransientGeometry

            Dim ln As Inventor.Line = TryCast(axisProxy.Geometry, Inventor.Line)
            If ln Is Nothing Then Return

            Dim p0 As Inventor.Point = ln.RootPoint
            Dim v As Inventor.Vector = ln.Direction
            If v Is Nothing Then Return

            Dim vv As Double = v.DotProduct(v)
            If vv <= 0 Then Return

            ' spawnPt를 현재 축 라인에 직교투영한 점 cp
            Dim w As Inventor.Vector = tg.CreateVector(spawnPt.X - p0.X, spawnPt.Y - p0.Y, spawnPt.Z - p0.Z)
            Dim t As Double = w.DotProduct(v) / vv

            Dim cp As Inventor.Point = tg.CreatePoint(p0.X + v.X * t, p0.Y + v.Y * t, p0.Z + v.Z * t)

            ' 축을 spawnPt로 옮기기 위한 translation = spawnPt - cp
            Dim dx As Double = spawnPt.X - cp.X
            Dim dy As Double = spawnPt.Y - cp.Y
            Dim dz As Double = spawnPt.Z - cp.Z

            ' 변위가 너무 작으면 스킵
            If Math.Abs(dx) < 0.000001 AndAlso Math.Abs(dy) < 0.000001 AndAlso Math.Abs(dz) < 0.000001 Then Return

            Dim m As Matrix = occ.Transformation
            Dim mm As Matrix = tg.CreateMatrix()
            mm.SetToIdentity()
            mm.Cell(1, 4) = dx
            mm.Cell(2, 4) = dy
            mm.Cell(3, 4) = dz

            ' 현재 변환 * 이동 변환
            m.TransformBy(mm)
            occ.Transformation = m

            Try
                app.ActiveView.Update()
            Catch
            End Try
        Catch
        End Try
    End Sub
    ' [NEW] Support Occ에서 "축(Line)에 가장 가까운 WorkPoint"를 찾아 Proxy 반환
    ' - spawnPoint 근처가 아니라 "축 위 점"을 찾는 게 핵심(오프셋/날아감 방지)
    Private Function GetSupportWorkPointProxyOnAxis(occ As ComponentOccurrence,
                                               axisProxy As WorkAxisProxy) As WorkPointProxy
        Try
            If occ Is Nothing OrElse axisProxy Is Nothing Then Return Nothing

            Dim app = Globals.g_inventorApplication
            If app Is Nothing Then Return Nothing

            Dim tg = app.TransientGeometry

            Dim partDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
            If partDef Is Nothing Then Return Nothing

            Dim ln As Inventor.Line = TryCast(axisProxy.Geometry, Inventor.Line)
            If ln Is Nothing Then
                ' 축 geometry가 Line이 아니면 fallback
                GoTo FALLBACK_WP1
            End If

            Dim p0 As Inventor.Point = ln.RootPoint
            Dim v As Inventor.Vector = ln.Direction
            If v Is Nothing OrElse v.Length <= 0 Then GoTo FALLBACK_WP1

            Dim bestWp As WorkPoint = Nothing
            Dim bestDist As Double = Double.MaxValue

            For Each wp As WorkPoint In partDef.WorkPoints
                If wp Is Nothing Then Continue For

                Dim lp As Inventor.Point = Nothing
                Try
                    lp = TryCast(wp.Point, Inventor.Point)
                Catch
                    lp = Nothing
                End Try
                If lp Is Nothing Then Continue For

                ' 로컬→전역
                Dim pg As Inventor.Point = tg.CreatePoint(lp.X, lp.Y, lp.Z)
                pg.TransformBy(occ.Transformation)

                ' 점-직선 거리 = |(pg-p0) x v| / |v|
                Dim w As Inventor.Vector = tg.CreateVector(pg.X - p0.X, pg.Y - p0.Y, pg.Z - p0.Z)
                Dim c As Inventor.Vector = w.CrossProduct(v)
                If c Is Nothing Then Continue For

                Dim d As Double = c.Length / v.Length

                If d < bestDist Then
                    bestDist = d
                    bestWp = wp
                End If
            Next

            If bestWp Is Nothing Then GoTo FALLBACK_WP1

            Dim prxObj As Object = Nothing
            occ.CreateGeometryProxy(bestWp, prxObj)
            Return TryCast(prxObj, WorkPointProxy)

FALLBACK_WP1:
            ' 마지막 fallback: WorkPoints(1)
            Try
                Dim wp1 As WorkPoint = partDef.WorkPoints.Item(1)
                Dim prxObj1 As Object = Nothing
                occ.CreateGeometryProxy(wp1, prxObj1)
                Return TryCast(prxObj1, WorkPointProxy)
            Catch
            End Try

            Return Nothing
        Catch
            Return Nothing
        End Try
    End Function
    ' [CHANGED] targetOcc에서 pick 기준으로 가장 가까운 "원통 FaceProxy"를 잡고,
    '          그 원통의 "중심축"을 Assembly WorkAxis로 만들어 WorkAxisProxy 반환
    ' - 변경 포인트:
    '   1) Definition.Face 루프 대신 GetCylinderFaceProxiesFromOccurrence(Proxy 기반) 사용
    '   2) 거리측정 실패/프록시 실패에 강함
    '   3) Cylinder 캐스팅 실패 시 fp.NativeObject(Face)로 한번 더 시도
    Private Function GetTargetCylinderAxisProxyNearPick(targetOcc As ComponentOccurrence,
                                                    pickPt As Inventor.Point) As WorkAxisProxy
        Try
            If Globals.g_inventorApplication Is Nothing Then Return Nothing
            If targetOcc Is Nothing Then Return Nothing

            Dim app = Globals.g_inventorApplication
            Dim tg = app.TransientGeometry

            ' 1) Assembly 컨텍스트에서 원통 FaceProxy 후보 수집
            Dim candidates As List(Of FaceProxy) = GetCylinderFaceProxiesFromOccurrence(targetOcc)
            If candidates Is Nothing OrElse candidates.Count = 0 Then Return Nothing

            ' 2) pickPt에 가장 가까운 원통 FaceProxy 선택(거리측정 실패는 스킵)
            Dim bestFp As FaceProxy = Nothing
            Dim bestDist As Double = Double.MaxValue

            If pickPt IsNot Nothing Then
                For Each fp As FaceProxy In candidates
                    If fp Is Nothing Then Continue For

                    Dim d As Double = Double.MaxValue
                    Try
                        d = app.MeasureTools.GetMinimumDistance(pickPt, fp)
                    Catch
                        Continue For
                    End Try

                    If d < bestDist Then
                        bestDist = d
                        bestFp = fp
                    End If
                Next
            End If

            ' pickPt가 없거나(혹은 전부 거리측정 실패)면 "반지름 최대"로 fallback
            If bestFp Is Nothing Then
                Dim bestR As Double = -1.0
                For Each fp As FaceProxy In candidates
                    If fp Is Nothing Then Continue For

                    Dim cyl0 As Cylinder = Nothing
                    Try
                        cyl0 = TryCast(fp.Geometry, Cylinder)
                    Catch
                        cyl0 = Nothing
                    End Try

                    If cyl0 Is Nothing Then
                        Try
                            Dim nf As Face = TryCast(fp.NativeObject, Face)
                            If nf IsNot Nothing Then cyl0 = TryCast(nf.Geometry, Cylinder)
                        Catch
                            cyl0 = Nothing
                        End Try
                    End If

                    If cyl0 Is Nothing Then Continue For

                    If cyl0.Radius > bestR Then
                        bestR = cyl0.Radius
                        bestFp = fp
                    End If
                Next
            End If

            If bestFp Is Nothing Then Return Nothing

            ' 3) 원통 기하에서 축 정보 추출 (Cylinder 확보)
            Dim cyl As Cylinder = Nothing
            Try
                cyl = TryCast(bestFp.Geometry, Cylinder)
            Catch
                cyl = Nothing
            End Try

            If cyl Is Nothing Then
                Try
                    Dim nativeFace As Face = TryCast(bestFp.NativeObject, Face)
                    If nativeFace IsNot Nothing Then
                        cyl = TryCast(nativeFace.Geometry, Cylinder)
                    End If
                Catch
                    cyl = Nothing
                End Try
            End If

            If cyl Is Nothing Then Return Nothing

            ' 4) 축 위 점(파트좌표) + 축 방향(파트좌표) → Assembly 좌표 변환
            Dim p0 As Inventor.Point = tg.CreatePoint(cyl.BasePoint.X, cyl.BasePoint.Y, cyl.BasePoint.Z)
            Dim v0 As Inventor.Vector = tg.CreateVector(cyl.AxisVector.X, cyl.AxisVector.Y, cyl.AxisVector.Z)

            p0.TransformBy(targetOcc.Transformation)
            v0.TransformBy(targetOcc.Transformation)

            If v0 Is Nothing OrElse v0.Length <= 0 Then Return Nothing

            Dim u0 As UnitVector = tg.CreateUnitVector(v0.X, v0.Y, v0.Z)

            ' 5) Assembly WorkAxis 생성 후 Proxy 반환
            Dim asm As AssemblyDocument = TryCast(app.ActiveDocument, AssemblyDocument)
            If asm Is Nothing Then Return Nothing

            Dim asmDef As AssemblyComponentDefinition = asm.ComponentDefinition

            Try
                If m_tempWorkAxis IsNot Nothing Then m_tempWorkAxis.Delete()
            Catch
            End Try
            m_tempWorkAxis = Nothing

            m_tempWorkAxis = asmDef.WorkAxes.AddFixed(p0, u0, False)

            Dim axisProxyObj As Object = Nothing
            asmDef.CreateGeometryProxy(m_tempWorkAxis, axisProxyObj)

            Return TryCast(axisProxyObj, WorkAxisProxy)

        Catch
            Return Nothing
        End Try
    End Function

#End Region

End Class
