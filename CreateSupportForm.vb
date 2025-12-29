' ========================= Part 1/2 =========================
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

            ' Distance는 기본 0으로(원하면 0.00 말고 기본값으로 바꿔도 됨)
            If Distance IsNot Nothing Then Distance.Text = "0.00"

            ' 옵션/토글 버튼 초기화
            ResetAllButtons()
            UpdateOptionButtons()
            InitSingleMultiToggle()

            ' 미리보기 이미지 초기화(기존 로고 규칙 그대로)
            If Preview IsNot Nothing Then
                Preview.Image = CreateSupport.Resource1.logo
            End If

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
        If mat <> "STS" Then
            UpdateFilename()
            Return
        End If

        ' STS: SF -> SS -> SR -> SL -> SF ...
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

        btype.Enabled = enableMain
        nextbt.Enabled = enableMain

        If enableMain Then
            btype.BackColor = cBtnNormal
            nextbt.BackColor = cBtnNormal
        Else
            btype.Checked = False
            btype.BackColor = Draw.Color.FromArgb(60, 60, 60)
            nextbt.BackColor = Draw.Color.FromArgb(60, 60, 60)
        End If

        If saddle.Checked Then
            Preview.Image = CreateSupport.Resource1.saddleimage
        ElseIf grip.Checked Then
            Preview.Image = CreateSupport.Resource1.gripimage
        Else
            Preview.Image = CreateSupport.Resource1.logo
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

    ' Mate/Flush 면 하이라이트(확정 전까지 유지)
    Private m_hlBaseFace As HighlightSet = Nothing
    Private m_hlTargetFace As HighlightSet = Nothing

    Private m_baseFace As Face
    Private m_targetFace As Face
    Private m_isFlushMode As Boolean = False

    Private isObjectSelectionActive As Boolean = False

    Private m_selectedComponent As ComponentOccurrence
    Private m_centerPoint As Inventor.Point
    Private m_supportFace As Face

    Private m_supportCompHighlight As HighlightSet
    Private m_supportFaceHighlight As HighlightSet

    ' [파일명용] STS LegType 코드 (기본 SF)
    Private m_legCode As String = "SF"

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


    Private previewImages As List(Of Draw.Image)
    Private currentPreviewIndex As Integer = 0



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

        Try
            Dim bmp As New Draw.Bitmap(CreateSupport.Resource1.logo)
            bmp.MakeTransparent(Draw.Color.Black)
            logoBox.Image = bmp
            logoBox.SizeMode = PictureBoxSizeMode.Zoom
            logoBox.Parent = pnlTitle
            lblTitle.Parent = pnlTitle
        Catch
        End Try

        previewImages = New List(Of Draw.Image) From {
            CreateSupport.Resource1.test1,
            CreateSupport.Resource1.test2,
            CreateSupport.Resource1.test3,
            CreateSupport.Resource1.test4
        }
        If previewImages.Count > 0 Then Preview.Image = previewImages(0)

        AddHandler Distance.MouseWheel, AddressOf Distance_MouseWheel
        AddHandler Distance.TextChanged, AddressOf Distance_TextChanged

        ' [NEW] 드래그로 창 이동 (원하는 영역을 잡고 끌기)
        AddHandler Me.MouseDown, AddressOf DragArea_MouseDown
        AddHandler Me.MouseMove, AddressOf DragArea_MouseMove
        AddHandler Me.MouseUp, AddressOf DragArea_MouseUp

        ' 타이틀/상단 패널이 있으면 같이 연결(있을 때만)
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

    Private Sub ClearPreviewAndHighlights()
        Try
            m_blueHighlight?.Delete()
        Catch
        End Try
        m_blueHighlight = Nothing

        Try
            m_greenHighlight?.Delete()
        Catch
        End Try
        m_greenHighlight = Nothing

        Try
            m_supportCompHighlight?.Delete()
        Catch
        End Try
        m_supportCompHighlight = Nothing

        Try
            m_supportFaceHighlight?.Delete()
        Catch
        End Try
        m_supportFaceHighlight = Nothing
        ' [NEW] 임시 배치된 Support Occurrence 삭제 (Cancel / FormClosed 시)
        Try
            If m_supportOcc IsNot Nothing Then m_supportOcc.Delete()
        Catch
        End Try
        m_supportOcc = Nothing

        ' [NEW] 임시 배치된 SAD Occurrence 삭제
        Try
            If m_sadOcc IsNot Nothing Then m_sadOcc.Delete()
        Catch
        End Try
        m_sadOcc = Nothing

        ' [NEW] 축대축 임시 구속 2개 삭제
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

        ' Mate/Flush 임시 구속 삭제
        Try
            If m_tempMateFlushConstraint IsNot Nothing Then m_tempMateFlushConstraint.Delete()
        Catch
        End Try
        m_tempMateFlushConstraint = Nothing

        ' Mate/Flush 면 하이라이트 삭제
        Try
            m_hlBaseFace?.Delete()
        Catch
        End Try
        m_hlBaseFace = Nothing

        Try
            m_hlTargetFace?.Delete()
        Catch
        End Try
        m_hlTargetFace = Nothing

        ' [NEW] 임시 WorkAxis 삭제
        Try
            If m_tempWorkAxis IsNot Nothing Then
                m_tempWorkAxis.Delete()
            End If
        Catch
        End Try
        m_tempWorkAxis = Nothing

        m_baseFace = Nothing
        m_targetFace = Nothing

        m_selectedComponent = Nothing
        m_centerPoint = Nothing
        m_supportFace = Nothing

        isObjectSelectionActive = False
    End Sub
    ' ===================== [NEW] 임시 생성 파일 삭제 유틸 =====================
    Private Sub DeleteTempCreatedFiles()
        Try
            For Each p In m_tempCreatedFiles
                If IO.File.Exists(p) Then
                    IO.File.Delete(p)
                End If
            Next
        Catch
        End Try
        m_tempCreatedFiles.Clear()
    End Sub
    ' [NEW] Cancel/X/닫기 시: 임시배치(occ) + 임시구속 + 임시축 + 하이라이트까지 전부 정리
    Private Sub ClearAll_OnClose()

        ' 임시 배치 Support Occ 삭제
        Try
            If m_supportOcc IsNot Nothing Then m_supportOcc.Delete()
        Catch
        End Try
        m_supportOcc = Nothing

        ' 임시 배치 SAD Occ 삭제
        Try
            If m_sadOcc IsNot Nothing Then m_sadOcc.Delete()
        Catch
        End Try
        m_sadOcc = Nothing

        ' 축대축 임시 구속 삭제 (Support / SAD)
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

        ' Mate/Flush 임시 구속 삭제
        Try
            If m_tempMateFlushConstraint IsNot Nothing Then m_tempMateFlushConstraint.Delete()
        Catch
        End Try
        m_tempMateFlushConstraint = Nothing

        ' 임시 WorkAxis 삭제
        Try
            If m_tempWorkAxis IsNot Nothing Then m_tempWorkAxis.Delete()
        Catch
        End Try
        m_tempWorkAxis = Nothing

        ' (요구사항) 하이라이트는 Cancel/OK/Apply/X 때만 제거
        ClearMateFlushFaceHighlights()

        ' 기타 하이라이트/상태 정리
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
        ' [NEW] Cancel/X/닫기 시: 임시 생성 파일 삭제
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

        ' 1) 확정(Commit): 임시 구속/Occ는 Delete하지 않고 "변수만 해제"해서 남김
        m_tempAxisConstraint_Support = Nothing
        m_tempAxisConstraint_SAD = Nothing
        m_tempMateFlushConstraint = Nothing

        m_supportOcc = Nothing
        m_sadOcc = Nothing
        m_currentSadSourcePath = ""
        ' Mate/Flush 면 하이라이트 제거(요구사항)
        ClearMateFlushFaceHighlights()

        ' Apply에서만 저장(경고창 제거 핵심)
        SaveTempCreatedFiles()

        ' Apply는 "생성 유지"
        m_tempCreatedFiles.Clear()

        ' 2) 다음 선택 준비: Inventor Pick 초기화
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

        ' 4) UI 초기 상태로 리셋(텍스트 박스 포함)
        ResetUIToInitialState()

        ' 5) 바로 다음 TUBE/PIPE 선택 루프로 복귀
        isObjectSelectionActive = True
        Me.BeginInvoke(New Action(AddressOf RunAnalysisSequence))
    End Sub

    ' [NEW] OK: 임시 배치를 확정하고 창 닫기
    Private Sub ok_Click(sender As Object, e As EventArgs) Handles ok.Click
        ' 배치된 게 있다면 확정
        If m_supportOcc IsNot Nothing OrElse m_sadOcc IsNot Nothing Then
            m_tempAxisConstraint_Support = Nothing
            m_tempAxisConstraint_SAD = Nothing
            m_tempMateFlushConstraint = Nothing

            m_supportOcc = Nothing
            m_sadOcc = Nothing
            m_currentSadSourcePath = ""
        End If

        ClearMateFlushFaceHighlights()
        SaveTempCreatedFiles()
        m_tempCreatedFiles.Clear()

        ' 창 닫기
        Me.Close()
    End Sub

    ' [NEW] 실제 배치 및 미리보기 수행 (기존 Apply 로직 이동 + 재배치 시 기존 삭제 기능 추가)
    Private Sub RunPreviewPlacement()

        ' 기존 미리보기 Occ가 있으면 Delete보다 Replace를 우선 사용
        Dim needNewOcc As Boolean = (m_supportOcc Is Nothing)

        ' 2) Support 파일 찾기
        m_resolvedSupportPath = FindSupportFile()
        If String.IsNullOrWhiteSpace(m_resolvedSupportPath) Then
            MessageBox.Show(
            "조건에 맞는 Support 파일을 찾지 못했습니다." & vbCrLf &
            "폴더: " & GetSupportFolderPath() & vbCrLf &
            "키: STS / " & ConvertSingleMultiToCode() & " / " & GetLegCode() & " / " & ConvertBoltToCode(),
            "찾기 실패"
        )
            Return
        End If

        ' 2-1) STS일 때만 SAD 파일 추가로 찾기 (STS_SUPPORT 폴더 바로 아래)
        m_resolvedSadPath = ""
        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()

        If mat = "STS" Then
            Dim stsFolder As String = IO.Path.Combine(SUPPORT_BASE_PATH, "STS_SUPPORT")
            m_resolvedSadPath = FindStsSadFile(stsFolder)

            If String.IsNullOrWhiteSpace(m_resolvedSadPath) Then
                MessageBox.Show(
                "SAD 포함 Support 파일을 찾지 못했습니다." & vbCrLf &
                "폴더: " & stsFolder & vbCrLf &
                "키: STS / " & ConvertSingleMultiToCode() & " / " & ConvertTypeToCode(ttext.Text) & " / SAD / " & ConvertSizeToCode(stext.Text),
                "SAD 찾기 실패"
            )
                ' SAD는 선택 기능 → 못 찾으면 그냥 Support만 진행
                m_resolvedSadPath = ""
            End If
        End If

        ' 3) 어셈블리 배치(기존 Support)
        '    ★ 파일명은 splength.Text 기반으로 이미 filename.Text가 만들어져 있음 (여기는 그대로 사용)
        Dim newFileName As String = If(filename.Text, "").Trim()
        If String.IsNullOrWhiteSpace(newFileName) Then Return

        Dim asm As AssemblyDocument = TryCast(Globals.g_inventorApplication.ActiveDocument, AssemblyDocument)
        If asm Is Nothing Then Return

        Dim asmFolder As String = IO.Path.GetDirectoryName(asm.FullFileName)

        ' '.' 은 파일명에서 P로 치환해서 실제 저장명과 일치시킴
        Dim expectedName As String = newFileName.Replace(".", "P") & ".ipt"
        Dim dstPathExpected As String = IO.Path.Combine(asmFolder, expectedName)

        ' ★ 이미 존재하면 재사용(덮어쓰기/경고창 방지)
        Dim existedAlready As Boolean = IO.File.Exists(dstPathExpected)

        ' ★ 있으면 그걸 쓰고, 없으면 복사 생성
        Dim copiedOrExistingPath As String = ResolveOrCopySupportToAsmFolder(m_resolvedSupportPath, newFileName)
        If String.IsNullOrWhiteSpace(copiedOrExistingPath) Then Return

        ' ★ "새로 생성된 경우"에만 임시 생성 파일 목록에 넣는다
        If Not existedAlready Then
            If Not m_tempCreatedFiles.Contains(copiedOrExistingPath) Then
                m_tempCreatedFiles.Add(copiedOrExistingPath)
            End If
        End If

        ' ★ 배치
        If m_supportOcc Is Nothing Then
            ' ★ 우선순위: 클릭 지점 → 없으면 centerPoint
            Dim spawnPt As Inventor.Point = If(m_lastPickPoint, m_centerPoint)
            m_supportOcc = PlaceSupportOccurrence(copiedOrExistingPath, spawnPt)
            If m_supportOcc Is Nothing Then Return
        Else
            Try
                ' Replace로 대치 (구속/위치 유지 기대)
                m_supportOcc.Replace(copiedOrExistingPath, True)
                Globals.g_inventorApplication.ActiveView.Update()
            Catch
                ' Replace 실패 시만 삭제 후 재배치로 폴백
                Try : m_supportOcc.Delete() : Catch : End Try
                Dim spawnPt As Inventor.Point = If(m_lastPickPoint, m_centerPoint)
                m_supportOcc = PlaceSupportOccurrence(copiedOrExistingPath, spawnPt)
                If m_supportOcc Is Nothing Then Return
            End Try
        End If

        ' =====================================================
        ' 4) 파라미터 값 계산
        '    - SD : size offset (기존처럼)
        '    - SPL: ★ 모델링 파라미터는 length.Text(원본거리) 기준
        ' =====================================================
        Dim sdVal As Double = GetSizeOffset(stext.Text)

        ' =====================================================
        ' ★ SPL(모델링)은 "length.Text" (원본거리)에서만 가져온다
        '    - filename은 splength.Text 기반(기존 유지)
        ' =====================================================
        Dim rawLenText As String = If(length.Text, "").Trim()
        Dim splVal As Double = 0

        ' 한국/영문 소수점 혼용 대비
        If Not Double.TryParse(rawLenText,
                       Globalization.NumberStyles.Any,
                       Globalization.CultureInfo.InvariantCulture,
                       splVal) Then
            Double.TryParse(rawLenText, splVal)
        End If

        ' ★ SPL은 "length.Text" 기준이므로, 파일이 이미 있어도 항상 갱신해야 함
        UpdateSupportParameters(m_supportOcc, sdVal, splVal)

        ' 3-1) SAD 배치(STS + SAD 파일이 있을 때만)
        '      - 핵심: "경로가 바뀔 때만" Replace/재배치
        If mat = "STS" AndAlso Not String.IsNullOrWhiteSpace(m_resolvedSadPath) Then

            Dim desiredSadPath As String = m_resolvedSadPath.Trim()

            ' (1) 이미 SAD가 있고, 같은 파일이면: 아무 것도 하지 않는다(유지)
            If m_sadOcc IsNot Nothing AndAlso
       Not String.IsNullOrWhiteSpace(m_currentSadSourcePath) AndAlso
       String.Equals(m_currentSadSourcePath, desiredSadPath, StringComparison.OrdinalIgnoreCase) Then
                ' 그대로 유지
            Else
                ' (2) SAD가 없거나, 경로가 바뀐 경우에만 업데이트
                If m_sadOcc Is Nothing Then
                    Dim spawnPtSad As Inventor.Point = If(m_lastPickPoint, m_centerPoint)
                    m_sadOcc = PlaceSupportOccurrence(desiredSadPath, spawnPtSad)
                    If m_sadOcc Is Nothing Then Return
                Else
                    ' 가능하면 Replace 우선(위치/구속 유지 기대)
                    Try
                        m_sadOcc.Replace(desiredSadPath, True)
                        Globals.g_inventorApplication.ActiveView.Update()
                    Catch
                        ' Replace 실패 시만 삭제 후 재배치
                        Try : m_sadOcc.Delete() : Catch : End Try
                        Dim spawnPtSad As Inventor.Point = If(m_lastPickPoint, m_centerPoint)
                        m_sadOcc = PlaceSupportOccurrence(desiredSadPath, spawnPtSad)
                        If m_sadOcc Is Nothing Then Return
                    End Try
                End If

                ' 현재 적용된 SAD 경로 갱신
                m_currentSadSourcePath = desiredSadPath
            End If

        Else
            ' STS가 아니거나 SAD 경로를 못 찾았으면: SAD 유지/삭제 정책은 선택
            ' 지금 요구사항(옵션 바꿀 때 불필요한 갱신 금지) 기준으로는 "그냥 둠"이 안전함.
            ' 필요하면 여기서 삭제하도록 바꿀 수 있음.
        End If

        ' 6) 축대축(또는 중심축) 임시구속 적용
        If m_selectedComponent Is Nothing Then
            ' 튜브가 선택되지 않았다면 배치만 하고 종료
            Return
        End If

        ' 6-1) 기존 Support 축 구속
        If Not ApplyTempAxisToAxisPreview(m_supportOcc, m_selectedComponent, m_tempAxisConstraint_Support, False) Then
            ' 실패 메시지는 내부에서 처리
        Else
            Globals.g_inventorApplication.StatusBarText = "Support(기본) 미리보기 배치 완료"
        End If

        ' 6-2) SAD 축 구속 (Axis_Z 이름 우선)
        If mat = "STS" AndAlso m_sadOcc IsNot Nothing Then
            If Not ApplyTempAxisToAxisPreview(m_sadOcc, m_selectedComponent, m_tempAxisConstraint_SAD, True) Then
                ' 실패 메시지는 내부에서 처리
            Else
                Globals.g_inventorApplication.StatusBarText = "Support(기본) + SAD 미리보기 배치 완료"
            End If
        End If

    End Sub

    ' =========================
    ' [NEW] 같은 파일명 있으면 재사용, 없으면 새로 복사
    ' =========================
    Private Function ResolveOrCopySupportToAsmFolder(srcLibPath As String, newFileName As String) As String
        Dim asm As AssemblyDocument = TryCast(Globals.g_inventorApplication.ActiveDocument, AssemblyDocument)
        If asm Is Nothing Then Return ""
        If String.IsNullOrWhiteSpace(srcLibPath) OrElse Not IO.File.Exists(srcLibPath) Then Return ""
        If String.IsNullOrWhiteSpace(newFileName) Then Return ""

        newFileName = newFileName.Replace(".", "P")

        Dim asmFolder As String = IO.Path.GetDirectoryName(asm.FullFileName)
        Dim dstPath As String = IO.Path.Combine(asmFolder, newFileName & ".ipt")

        ' ★ 이미 있으면: 덮어쓰기 금지(경고창 원인 제거) → 그대로 재사용
        If IO.File.Exists(dstPath) Then
            Return dstPath
        End If

        ' 없으면: 기존 로직대로 복사(생성)
        Try
            IO.File.Copy(srcLibPath, dstPath, False)
            Try : IO.File.SetAttributes(dstPath, IO.FileAttributes.Normal) : Catch : End Try
            Return dstPath
        Catch ex As Exception
            MessageBox.Show("파일 복사 실패: " & ex.Message)
            Return ""
        End Try
    End Function

    ' [NEW] et1=Mate / et2=Flush
    Private Sub et1_CheckedChanged(sender As Object, e As EventArgs) Handles et1.CheckedChanged
        If et1 Is Nothing Then Return
        If Not et1.Checked Then Return

        ' Mate 모드
        m_isFlushMode = False
        If et2 IsNot Nothing Then et2.Checked = False

        ' [CHANGED] 이미 면 2개가 있으면 재선택 없이 바로 전환
        If m_baseFace IsNot Nothing AndAlso m_targetFace IsNot Nothing Then
            RebuildTempMateFlushConstraint()
        Else
            Me.BeginInvoke(New Action(AddressOf StartMateFlushPickSequence))
        End If

    End Sub

    Private Sub et2_CheckedChanged(sender As Object, e As EventArgs) Handles et2.CheckedChanged
        If et2 Is Nothing Then Return
        If Not et2.Checked Then Return

        ' Flush 모드
        m_isFlushMode = True
        If et1 IsNot Nothing Then et1.Checked = False

        ' 면 2개 다시 선택해서 미리보기 구속 생성
        ' [CHANGED] 이미 면 2개가 있으면 재선택 없이 바로 전환
        If m_baseFace IsNot Nothing AndAlso m_targetFace IsNot Nothing Then
            RebuildTempMateFlushConstraint()
        Else
            Me.BeginInvoke(New Action(AddressOf StartMateFlushPickSequence))
        End If

    End Sub
    ' [NEW] Mate/Flush 면선택 → 임시 구속 생성
    Private Sub StartMateFlushPickSequence()
        If Globals.g_inventorApplication Is Nothing Then Return

        ' Mate/Flush 임시구속만 제거 (중심축 구속은 유지!)
        Try
            If m_tempMateFlushConstraint IsNot Nothing Then m_tempMateFlushConstraint.Delete()
        Catch
        End Try
        m_tempMateFlushConstraint = Nothing

        ' [CHANGED] "새로 선택할 때만" 초기화해야 함
        ' 기존 면이 이미 잡혀있으면 유지하고, StartMateFlushPickSequence가 호출된 경우는
        ' (즉, 면이 없어서) 새로 뽑는 상황이므로 여기서 초기화해도 OK

        m_baseFace = Nothing
        m_targetFace = Nothing

        ' Inventor로 포커스 + ESC로 Pick 초기화
        Try
            SetForegroundWindow(New IntPtr(Globals.g_inventorApplication.MainFrameHWND))
            SendKeys.SendWait("{ESC}")
        Catch
        End Try

        Try
            ' 1) Base Face
            Dim f1 = Globals.g_inventorApplication.CommandManager.Pick(
            SelectionFilterEnum.kPartFaceFilter,
            "기준 면(Base Face)을 선택하세요")

            If f1 Is Nothing OrElse Not TypeOf f1 Is Face Then Return
            m_baseFace = CType(f1, Face)
            HighlightMateFlushFace(m_baseFace, True)

            ' 2) Target Face
            Dim f2 = Globals.g_inventorApplication.CommandManager.Pick(
            SelectionFilterEnum.kPartFaceFilter,
            "맞출 면(Target Face)을 선택하세요")

            If f2 Is Nothing OrElse Not TypeOf f2 Is Face Then Return
            m_targetFace = CType(f2, Face)
            HighlightMateFlushFace(m_targetFace, False)

            ' 3) 구속 생성 + 현재 Distance 값 반영
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

        ' Mate/Flush 임시구속만 제거
        Try
            If m_tempMateFlushConstraint IsNot Nothing Then m_tempMateFlushConstraint.Delete()
        Catch
        End Try
        m_tempMateFlushConstraint = Nothing

        ' Distance(mm) → Inventor DB length로 변환해서 offset에 사용
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

            ' MateConstraint/FlushConstraint 둘 다 Offset을 가짐
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
            ' 숫자 입력 중(예: "-" / "." / 빈칸)엔 조용히 무시
        End Try
    End Sub

    ' [NEW] Distance.Text(mm) → DB length 변환
    Private Function GetOffsetDbFromDistanceText(uom As UnitsOfMeasure) As Double
        Dim s As String = If(Distance.Text, "").Trim()
        If s = "" Then Return 0

        Dim v As Double = 0
        If Not Double.TryParse(s, Globalization.NumberStyles.Any,
                           Globalization.CultureInfo.InvariantCulture, v) Then
            ' 한국식 입력 대비(콤마/소수)
            Double.TryParse(s, v)
        End If

        ' mm 입력값을 "mm"로 파싱 → DB length 반환
        Return uom.GetValueFromExpression(v.ToString("0.###", Globalization.CultureInfo.InvariantCulture) & " mm",
                                     UnitsTypeEnum.kMillimeterLengthUnits)
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
                    m_hlBaseFace.Color = app.TransientObjects.CreateColor(0, 191, 255) ' 파랑
                End If
                hs = m_hlBaseFace
            Else
                If m_hlTargetFace Is Nothing Then
                    m_hlTargetFace = doc.CreateHighlightSet()
                    m_hlTargetFace.Color = app.TransientObjects.CreateColor(120, 255, 120) ' 초록
                End If
                hs = m_hlTargetFace
            End If

            hs.Clear()
            hs.AddItem(faceObj)

            app.ActiveView.Update()
        Catch
        End Try
    End Sub
    ' [NEW] X 클릭 = 취소와 동일 (임시 배치/구속/축/하이라이트 정리 후 닫기)
    Private Sub lblClose_Click(sender As Object, e As EventArgs)
        ClearAll_OnClose()
        Me.Close()
    End Sub

    ' [NEW] X 호버 색상
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
                    PickOccurrenceWithPoint("SUPPORT 생성할 대상을 선택하세요 (ESC 종료)", pPick)

                oSelect = occPicked

                ' [NEW] 클릭 지점 저장(없으면 Nothing)
                m_lastPickPoint = pPick

            Catch
                Exit While
            End Try

            If oSelect Is Nothing OrElse Not TypeOf oSelect Is ComponentOccurrence Then Exit While

            Dim comp As ComponentOccurrence = CType(oSelect, ComponentOccurrence)
            Dim fileName As String = comp.Definition.Document.DisplayName.ToUpper()

            selHL?.Clear()
            selHL?.AddItem(comp)

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

        selHL?.Delete()
    End Sub

    ' [REPLACE] 면 선택 후 자동으로 길이 계산 및 "미리보기 배치" 실행
    Private Sub RunSupportFaceSelection()
        If m_centerPoint Is Nothing Then Return

        Try
            SetForegroundWindow(New IntPtr(Globals.g_inventorApplication.MainFrameHWND))
            Dim face = Globals.g_inventorApplication.CommandManager.Pick(
                SelectionFilterEnum.kPartFaceFilter,
                "지지할 바닥면 선택")

            If face Is Nothing Then Return

            ' Assembly에서 Pick하면 FaceProxy일 수 있어서 Object로 처리 후 Face로 캐스팅 시도
            Dim pickedObj As Object = face

            Dim realFace As Face = TryCast(pickedObj, Face)
            If realFace Is Nothing Then
                ' FaceProxy도 Face처럼 취급되지만, 안전하게 TryCast 한번 더
                realFace = TryCast(pickedObj, Face)
            End If
            If realFace Is Nothing Then Return

            m_supportFace = realFace

            ' [NEW] 지지면이 속한 파트의 Inventor "재질(Material)"을 읽어서 mtext에 넣는다
            Dim matCode As String = GetMaterialFromPickedFace(pickedObj)
            If Not String.IsNullOrWhiteSpace(matCode) Then
                mtext.Text = matCode   ' 예: "PVC", "STS"
            End If

            ' 재질이 바뀌면 파일명/검색키도 달라지니까 즉시 갱신
            UpdateOptionButtons()
            UpdateFilename()

            Dim distCm =
    Globals.g_inventorApplication.MeasureTools.GetMinimumDistance(m_centerPoint, m_supportFace)

            length.Text = (distCm * 10).ToString("0.00")
            UpdateSupportLength()
            UpdateFilename()

            ' [핵심 추가] 면 선택이 끝나면 즉시 Support를 불러와 임시 배치(Preview)
            Me.BeginInvoke(New Action(AddressOf RunPreviewPlacement))

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub



#End Region

#Region "J. SUPPORT 파일 탐색 규칙(기존 FindSupportFile)"
    ' [CHANGED] 재질명에서 괄호 제거 후,
    ' 포함 여부로 STS/PVC를 뽑아낸다 (STS316, STS304-HL, PVC-WHITE 등 대응)
    Private Function NormalizeMaterialName(rawName As String) As String
        Dim s As String = If(rawName, "").Trim()
        If s = "" Then Return ""


        ' 괄호(옵션명) 제거: "PVC(WHITE)" -> "PVC"
        Dim idx As Integer = s.IndexOf("("c)
        If idx >= 0 Then s = s.Substring(0, idx)

        s = s.Trim().ToUpper()

        ' 핵심: "포함" 기준으로 코드화
        If s.Contains("STS") OrElse s.Contains("SUS") OrElse s.Contains("STAINLESS") Then Return "STS"
        If s.Contains("PVC") Then Return "PVC"

        ' 둘 다 아니면: 일단 원문(괄호 제거된 재질명) 반환
        Return s
    End Function

    ' [NEW] 어셈블리에서 선택한 Face(대부분 FaceProxy)로부터 "해당 파트의 Material" 이름 가져오기
    Private Function GetMaterialFromPickedFace(faceObj As Object) As String
        Try
            If faceObj Is Nothing Then Return ""

            Dim matName As String = ""

            ' Assembly에서 Pick한 Face는 보통 FaceProxy로 들어옴
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
                ' Part 문서에서 직접 Face를 잡는 경우 대비
                Dim f As Face = TryCast(faceObj, Face)
                If f IsNot Nothing Then
                    Dim doc As Document = TryCast(f.Parent.Parent.Parent, Document) ' 안전하지 않을 수 있어 Try로 감쌈
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
        "O:\\System_BU_Vault\\00. 구매품 Library\\support\\"

    Private Function GetSupportFolderPath() As String
        Dim style =
            If(saddle.Checked, "Saddle",
               If(grip.Checked, "Grip", "Basic"))

        Select Case mtext.Text
            Case "PVC" : Return IO.Path.Combine(SUPPORT_BASE_PATH, "PVC_SUPPORT", style)
            Case "STS" : Return IO.Path.Combine(SUPPORT_BASE_PATH, "STS_SUPPORT")
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

        ' PVC는 나중에
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

            ' ★ spawnPoint가 있으면 그 위치에서 스폰
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

    ' [NEW] Support/SAD 파일의 사용자 매개변수 수정
    ' SD  : 사이즈 보정값(mm)
    ' SPL : 길이(mm)
    ' [CHANGED] Support/SAD 파일의 사용자 매개변수 수정
    ' - UserParameters에는 NameIsUsed가 없음 → Item("SD") 접근을 Try로 처리
    ' - 단위 꼬임 방지: Value(cm) 대신 Expression("mm")로 세팅
    ' ===================== existing(UpdateSupportParameters) → changed(전체 교체) =====================
    ' ===================== UpdateSupportParameters : 전체 교체 =====================
    Private Sub UpdateSupportParameters(occ As ComponentOccurrence, sdMm As Double, splMm As Double, Optional doSave As Boolean = False)
        If occ Is Nothing Then Return

        Dim partDef As PartComponentDefinition = TryCast(occ.Definition, PartComponentDefinition)
        If partDef Is Nothing Then Return

        Dim pdoc As PartDocument = TryCast(partDef.Document, PartDocument)
        If pdoc Is Nothing Then Return

        Dim sdExpr As String = sdMm.ToString("0.###", Globalization.CultureInfo.InvariantCulture) & " mm"
        Dim splExpr As String = splMm.ToString("0.###", Globalization.CultureInfo.InvariantCulture) & " mm"

        Dim okSD As Boolean = False
        Dim okSPL As Boolean = False

        ' 1) UserParameters 먼저
        Try
            Dim ups As UserParameters = partDef.Parameters.UserParameters
            Try : ups.Item("SD").Expression = sdExpr : okSD = True : Catch : End Try
            Try : ups.Item("SPL").Expression = splExpr : okSPL = True : Catch : End Try
        Catch
        End Try

        ' 2) 안 잡히면 ModelParameters에서 찾기
        If (Not okSD) OrElse (Not okSPL) Then
            Try
                Dim mps As ModelParameters = partDef.Parameters.ModelParameters
                For Each p As Parameter In mps
                    If p Is Nothing OrElse p.Name Is Nothing Then Continue For
                    Dim nm As String = p.Name.Trim().ToUpper()

                    If (Not okSD) AndAlso nm = "SD" Then
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

        ' 3) 실제로 들어간 Expression을 다시 읽어서 상태바에 출력
        Try
            Dim readSD As String = ""
            Dim readSPL As String = ""

            Try : readSD = partDef.Parameters.Item("SD").Expression : Catch : End Try
            Try : readSPL = partDef.Parameters.Item("SPL").Expression : Catch : End Try

            Globals.g_inventorApplication.StatusBarText =
            $"[SupportParam] SD={(If(okSD, "OK", "NG"))}({readSD})  SPL={(If(okSPL, "OK", "NG"))}({readSPL})  file={IO.Path.GetFileName(pdoc.FullFileName)}"
        Catch
        End Try

        ' 4) 업데이트만 (Preview에서는 저장 금지)
        Try : pdoc.Update2(True) : Catch : End Try

        ' 저장은 외부(Apply/OK)에서만 한다

    End Sub

    ' [NEW] 토큰이 "구분자(_ 또는 -)로 둘러싸인 상태"로 들어있는지 검사
    Private Function HasToken(nameUpper As String, tokenUpper As String) As Boolean
        nameUpper = If(nameUpper, "").Trim().ToUpper()
        tokenUpper = If(tokenUpper, "").Trim().ToUpper()
        If nameUpper = "" OrElse tokenUpper = "" Then Return False

        ' 구분자 = 알파/숫자가 아닌 모든 문자 (_,-,공백 등)
        Dim pattern As String =
        "(^|[^A-Z0-9])" &
        System.Text.RegularExpressions.Regex.Escape(tokenUpper) &
        "([^A-Z0-9]|$)"

        Return System.Text.RegularExpressions.Regex.IsMatch(nameUpper, pattern)
    End Function

    ' [NEW] STS: "STS / S|M / SF~SL / Y|N" 만으로 파일 찾기 (사이즈/길이/TU/PI/LSP 전부 무시)
    Private Function FindStsFileByCoreKeys_IgnoreSizeLength(stsFolder As String) As String
        If Not IO.Directory.Exists(stsFolder) Then Return ""

        Dim sm As String = ConvertSingleMultiToCode()   ' S/M
        Dim leg As String = GetLegCode()                ' SF/SS/SR/SL
        Dim bolt As String = ConvertBoltToCode()        ' Y/N

        Dim files = IO.Directory.GetFiles(stsFolder, "*.ipt", IO.SearchOption.TopDirectoryOnly)
        If leg <> "SF" AndAlso leg <> "SS" AndAlso leg <> "SR" AndAlso leg <> "SL" Then
            Return ""
        End If
        ' 1순위: STS + S/M + LEG + BOLT (토큰 매칭)
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, sm) Then Continue For
            If Not HasToken(n, leg) Then Continue For
            If Not HasToken(n, bolt) Then Continue For

            Return f
        Next

        ' 2순위: STS + LEG + BOLT (S/M 없는 파일 대비)
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, leg) Then Continue For
            If Not HasToken(n, bolt) Then Continue For

            Return f
        Next

        ' 3순위: STS + LEG (최후)
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, leg) Then Continue For

            Return f
        Next

        Return ""
    End Function

    ' [NEW] STS: 길이(Lxx) 무시하고 토큰으로 파일 찾기
    Private Function FindStsFileByTokensIgnoreLength(stsFolder As String) As String
        Dim sm As String = ConvertSingleMultiToCode()            ' S/M
        Dim typ As String = ConvertTypeToCode(ttext.Text)        ' TU/PI
        Dim mdl As String = ConvertModelToCode(Model.Text)       ' LSP
        Dim sizeCode As String = ConvertSizeToCode(stext.Text)   ' 06B 등
        Dim leg As String = GetLegCode()                         ' SF/SS/SR/SL
        Dim bolt As String = ConvertBoltToCode()                 ' Y/N

        Dim files = IO.Directory.GetFiles(stsFolder, "*.ipt", IO.SearchOption.TopDirectoryOnly)

        ' 1순위: STS + S/M + TU/PI + LSP + SIZE + LEG + BOLT 모두 만족 (길이 무시)
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, sm) Then Continue For
            If Not HasToken(n, typ) Then Continue For
            If Not HasToken(n, mdl) Then Continue For
            If Not HasToken(n, sizeCode) Then Continue For
            If Not HasToken(n, leg) Then Continue For
            If Not HasToken(n, bolt) Then Continue For

            Return f
        Next

        ' 2순위(너가 말한 "핵심만"): STS + S/M + LEG + BOLT
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, sm) Then Continue For
            If Not HasToken(n, leg) Then Continue For
            If Not HasToken(n, bolt) Then Continue For

            Return f
        Next

        Return ""
    End Function

    ' [NEW] STS: 불러오기용 핵심조건만으로 파일 1개 선택
    ' 조건: STS / (S|M) / (SF|SS|SR|SL) / (Y|N)
    Private Function FindStsFileByCoreKeys() As String
        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()
        If mat <> "STS" Then Return ""

        Dim folder As String = IO.Path.Combine(SUPPORT_BASE_PATH, "STS_SUPPORT")
        If Not IO.Directory.Exists(folder) Then Return ""

        Dim sm As String = ConvertSingleMultiToCode()     ' "S" or "M"
        Dim leg As String = GetLegCode()                  ' "SF".."SL"
        Dim bolt As String = ConvertBoltToCode()          ' "Y" or "N"

        ' 우선순위 매칭 패턴(정확 매칭 -> 부분 매칭)
        ' 파일명 어딘가에 "-S-" 또는 "_S_" 같은 구분자가 있을 수도 있으니 유연하게 검사
        Dim files = IO.Directory.GetFiles(folder, "*.ipt", IO.SearchOption.TopDirectoryOnly)

        ' 1) 가장 엄격: STS + S/M + LEG + BOLT 모두 포함
        For Each f In files
            Dim n = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If n.Contains("STS") AndAlso n.Contains(sm) AndAlso n.Contains(leg) AndAlso n.Contains(bolt) Then
                Return f
            End If
        Next

        ' 2) STS + LEG + BOLT
        For Each f In files
            Dim n = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If n.Contains("STS") AndAlso n.Contains(leg) AndAlso n.Contains(bolt) Then
                Return f
            End If
        Next

        ' 3) STS + LEG
        For Each f In files
            Dim n = IO.Path.GetFileNameWithoutExtension(f).ToUpper()
            If n.Contains("STS") AndAlso n.Contains(leg) Then
                Return f
            End If
        Next

        Return ""
    End Function
    Private Function ApplyTempAxisToAxisPreview(supportOcc As ComponentOccurrence,
                                            targetOcc As ComponentOccurrence,
                                            ByRef constraintStore As AssemblyConstraint,
                                            Optional preferAxisZName As Boolean = False) As Boolean
        If Globals.g_inventorApplication Is Nothing Then Return False
        If supportOcc Is Nothing OrElse targetOcc Is Nothing Then Return False

        Dim doc As Document = Globals.g_inventorApplication.ActiveDocument
        If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Return False

        Dim asm As AssemblyDocument = CType(doc, AssemblyDocument)
        Dim asmDef As AssemblyComponentDefinition = asm.ComponentDefinition

        ' 0) 기존 해당 구속 제거(넘겨받은 constraintStore)
        Try
            If constraintStore IsNot Nothing Then constraintStore.Delete()
        Catch
        End Try
        constraintStore = Nothing

        ' 1) Support 고정 해제
        Try
            supportOcc.Grounded = False
        Catch
        End Try

        ' 2) Support 축 Proxy (SAD는 Axis_Z 이름 우선)
        Dim supportAxisProxy As WorkAxisProxy = Nothing
        If preferAxisZName Then
            supportAxisProxy = GetAxisZProxyPreferNamed(supportOcc, "AXIS_Z")
        Else
            ' 기존처럼 "원점 Z축"
            supportAxisProxy = GetSupportOriginZAxisProxy(supportOcc)
        End If

        If supportAxisProxy Is Nothing Then
            MessageBox.Show("Support/SAD 파일에서 축을 찾을 수 없습니다.", "오류")
            Return False
        End If

        ' 3) 타겟 원통면 Proxy
        Dim targetFaceProxy As FaceProxy = GetTargetCylinderFaceProxy(targetOcc)
        If targetFaceProxy Is Nothing Then
            MessageBox.Show("타겟 파트에서 원통면을 찾을 수 없습니다.", "형상 찾기 실패")
            Return False
        End If

        ' 4) Mate 구속 (kInferredLine)
        Try
            constraintStore = asmDef.Constraints.AddMateConstraint(
            supportAxisProxy,
            targetFaceProxy,
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

    ' [NEW] 타겟 파트 내부의 WorkAxis 중, 원통면 중심과 정확히 일치하는 축을 찾아 Proxy로 반환
    ' [NEW] 축이 있으면 축을 반환, 없으면 원통면을 반환 (자동 중심축 구속용)
    ' [NEW] 타겟 파트의 "가장 큰 원통면"을 찾아 어셈블리 레벨 Proxy로 반환
    Private Function GetTargetCylinderFaceProxy(occ As ComponentOccurrence) As FaceProxy
        Try
            If occ Is Nothing Then Return Nothing

            Dim bestFace As Face = Nothing
            Dim maxRadius As Double = -1.0

            ' 파트 내의 모든 면을 검사하여 가장 큰 원통면(메인 파이프)을 찾음
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

            ' [중요] CreateGeometryProxy를 통해 어셈블리 맥락의 객체(Proxy)로 변환해야 함
            ' 이걸 안 하면 "매개 변수가 틀립니다(E_INVALIDARG)" 오류 발생
            Dim proxyObj As Object = Nothing
            occ.CreateGeometryProxy(bestFace, proxyObj)

            Return TryCast(proxyObj, FaceProxy)
        Catch
            Return Nothing
        End Try
    End Function

    ' [NEW] Support의 "원점 Z축"을 정확히 가져오는 함수 (Item 3 = Z축)
    Private Function GetSupportOriginZAxisProxy(occ As ComponentOccurrence) As WorkAxisProxy
        Try
            If occ Is Nothing Then Return Nothing
            ' Inventor Part 파일의 WorkAxes 1=X, 2=Y, 3=Z (고정 불변)
            Dim zAxis As WorkAxis = occ.Definition.WorkAxes.Item(3)

            Dim proxyObj As Object = Nothing
            occ.CreateGeometryProxy(zAxis, proxyObj)
            Return TryCast(proxyObj, WorkAxisProxy)
        Catch
            Return Nothing
        End Try
    End Function

    ' [NEW] STS + (S|M) + (TU|PI) + "SAD" + (SIZE) 로 SAD 파일 찾기
    ' 예: STS-S-TU-SAD-06B  (구분자 _ 또는 - 기준 토큰 매칭)
    Private Function FindStsSadFile(stsFolder As String) As String
        If Not IO.Directory.Exists(stsFolder) Then Return ""

        Dim sm As String = ConvertSingleMultiToCode()           ' S/M
        Dim typ As String = ConvertTypeToCode(ttext.Text)       ' TU/PI
        Dim sizeCode As String = ConvertSizeToCode(stext.Text)  ' 06B 등

        Dim files = IO.Directory.GetFiles(stsFolder, "*.ipt", IO.SearchOption.TopDirectoryOnly)

        ' 1순위: STS + S/M + TU/PI + SAD + SIZE
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, sm) Then Continue For
            If Not HasToken(n, typ) Then Continue For
            If Not HasToken(n, "SAD") Then Continue For
            If Not HasToken(n, sizeCode) Then Continue For

            Return f
        Next

        ' 2순위: STS + TU/PI + SAD + SIZE (S/M 없는 파일 대비)
        For Each f In files
            Dim n As String = IO.Path.GetFileNameWithoutExtension(f).ToUpper()

            If Not HasToken(n, "STS") Then Continue For
            If Not HasToken(n, typ) Then Continue For
            If Not HasToken(n, "SAD") Then Continue For
            If Not HasToken(n, sizeCode) Then Continue For

            Return f
        Next

        Return ""
    End Function
    ' [NEW] Support/SAD 파트에서 Axis_Z(이름) 우선, 없으면 원점 Z축(Item(3))으로 Proxy 반환
    Private Function GetAxisZProxyPreferNamed(occ As ComponentOccurrence, Optional axisName As String = "AXIS_Z") As WorkAxisProxy
        Try
            If occ Is Nothing Then Return Nothing

            Dim wa As WorkAxis = Nothing

            ' 1) 이름으로 먼저 찾기
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

            ' 2) 없으면 원점 Z축(3)
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
    ' [NEW] 길이값을 파일명용 문자열로 변환
    ' 예: 27.5 → L27P5 / 30 → L30
    ' [CHANGED] 소수점 문화권(, / .) 혼용 입력에도 안정적으로 동작하도록 파싱 로직 강화
    Private Function ConvertLengthToFileToken(lenText As String) As String
        Dim raw As String = If(lenText, "").Trim()
        If raw = "" Then Return "L0"


        Dim v As Double = 0

        ' 1) "." 기준(인벤터/코드에서 Invariant로 뽑아낸 값) 먼저 시도
        If Not Double.TryParse(raw,
                           Globalization.NumberStyles.Any,
                           Globalization.CultureInfo.InvariantCulture,
                           v) Then
            ' 2) 현재 문화권(한국식 , 소수점 등)도 허용
            If Not Double.TryParse(raw, v) Then
                Return "L0"
            End If
        End If

        Dim s As String = v.ToString("0.###", Globalization.CultureInfo.InvariantCulture)
        s = s.Replace(".", "P")   ' 소수점 → P
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
        If btype IsNot Nothing AndAlso btype.Checked Then Return "Y"
        Return "N"
    End Function

    Private Function GetLegCode() As String
        Dim mat As String = If(mtext.Text, "").Trim().ToUpper()
        If mat = "STS" Then Return m_legCode
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
             ' ---------- PIPE ----------
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

        ' [CHANGED] 파일명/파싱 흔들림 방지: splength는 항상 Invariant(소수점 ".")로 출력
        splength.Text = Math.Max(0, baseLen - GetSizeOffset(stext.Text)).ToString("0.00", Globalization.CultureInfo.InvariantCulture)
        UpdateFilename()
    End Sub



#End Region

#Region "M. Geometry Helpers(외경/사이즈/오차)"


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
        Dim t As String = (typeText & "").Trim().ToUpper()

        If t = "TUBE" Then
            If IsInRange(dia, 25.4, 1.0) Then Return "1"""
            If IsInRange(dia, 19.05, 1.0) Then Return "3/4"""
            If IsInRange(dia, 12.7, 1.0) Then Return "1/2"""
            If IsInRange(dia, 9.53, 1.0) Then Return "3/8"""
            If IsInRange(dia, 6.35, 1.0) Then Return "1/4"""
            If IsInRange(dia, 4.76, 1.0) Then Return "3/16"""
            If IsInRange(dia, 3.18, 1.0) Then Return "1/8"""
            Return Math.Round(dia, 2).ToString() & "mm"
        End If

        If t = "PIPE" Then
            If IsInRange(dia, 60.33, 1.0) Then Return "50A"
            If IsInRange(dia, 48.26, 1.0) Then Return "40A"
            If IsInRange(dia, 33.4, 1.0) Then Return "25A"
            If IsInRange(dia, 26.67, 1.0) Then Return "20A"
            If IsInRange(dia, 21.34, 1.0) Then Return "15A"
            Return Math.Round(dia, 2).ToString() & "mm"
        End If

        Return ""
    End Function

    Private Function IsInRange(value As Double, target As Double, tol As Double) As Boolean
        Return Math.Abs(value - target) <= tol
    End Function
    ' [NEW] Occurrence + 클릭한 3D 포인트(ModelPosition)까지 얻는 Pick
    ' [NEW] Occurrence + 클릭한 3D 포인트(ModelPosition)까지 얻는 Pick
    Private Function PickOccurrenceWithPoint(prompt As String, ByRef pickedPt As Inventor.Point) As ComponentOccurrence
        pickedPt = Nothing
        If Globals.g_inventorApplication Is Nothing Then Return Nothing

        Dim app = Globals.g_inventorApplication
        Dim ie As InteractionEvents = Nothing
        Dim se As SelectEvents = Nothing

        Dim done As Boolean = False
        Dim pickedOcc As ComponentOccurrence = Nothing
        Dim pickedModelPos As Inventor.Point = Nothing

        Try
            ie = app.CommandManager.CreateInteractionEvents()
            se = ie.SelectEvents

            se.AddSelectionFilter(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter)
            se.SingleSelectEnabled = True

            Dim onSelectHandler As SelectEventsSink_OnSelectEventHandler =
            Sub(JustSelectedEntities As ObjectsEnumerator,
                SelectionDevice As SelectionDeviceEnum,
                ModelPosition As Inventor.Point,
                ViewPosition As Inventor.Point2d,
                View As Inventor.View)

                Try
                    If JustSelectedEntities Is Nothing OrElse JustSelectedEntities.Count < 1 Then Return
                    Dim occ = TryCast(JustSelectedEntities.Item(1), ComponentOccurrence)
                    If occ Is Nothing Then Return

                    pickedOcc = occ
                    pickedModelPos = ModelPosition
                    done = True
                    Try : ie.Stop() : Catch : End Try
                Catch
                End Try
            End Sub

            Dim onTerminateHandler As InteractionEventsSink_OnTerminateEventHandler =
            Sub()
                ' ESC / 우클릭 종료 등
                done = True
            End Sub

            AddHandler se.OnSelect, onSelectHandler
            AddHandler ie.OnTerminate, onTerminateHandler

            ie.StatusBarText = prompt
            ie.Start()

            Do While isFormRunning AndAlso Not done AndAlso ie IsNot Nothing
                System.Windows.Forms.Application.DoEvents()
            Loop

            Try
                RemoveHandler se.OnSelect, onSelectHandler
            Catch
            End Try

            Try
                RemoveHandler ie.OnTerminate, onTerminateHandler
            Catch
            End Try

        Catch
        Finally
            Try
                If ie IsNot Nothing Then ie.Stop()
            Catch
            End Try
        End Try

        pickedPt = pickedModelPos
        Return pickedOcc
    End Function

#End Region

End Class