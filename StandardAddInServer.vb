' ============================================================
' CreateSupport — Inventor Add-In 메인 클래스
' 기능 요약:
'   ① Inventor 실행 시 자동 로드 (Activate 실행)
'   ② HanyangENG 탭에 "HYE Support" 패널 + "Create Support" 버튼 추가
'   ③ 버튼 클릭 시 CreateSupportForm 창을 띄움
'      - 이미 떠 있으면 새로 만들지 않고 활성화만 함
' ============================================================

Imports System.Runtime.InteropServices
Imports Inventor
Imports stdole
Imports System.Windows.Forms
Imports WinForms = System.Windows.Forms
Imports Drawing = System.Drawing




Namespace CreateSupport

    <ProgIdAttribute("CreateSupport.StandardAddInServer"),
    GuidAttribute("E98684B1-DC81-4B1C-B0CA-F078700E855C")>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' 이벤트를 Handles로 직접 걸지 않고 AddHandler를 사용하므로
        ' WithEvents는 UIEvents에만 사용하고 버튼은 일반 필드로 둔다.
        Private WithEvents m_uiEvents As UserInterfaceEvents   ' 리본 리셋 이벤트용
        Private m_supportButton As ButtonDefinition            ' 버튼 정의 (이벤트는 AddHandler로 연결)


        ' 탭: HYE Piping / id_hyePiping
        Private Const TAB_DISPLAY As String = "HYE Piping"      ' 탭에 보일 이름
        Private Const TAB_INTERNAL As String = "id_hyePiping"   ' 리본 탭 InternalName

        ' 패널: HYE Support / id_hyeSupportPanel
        Private Const PANEL_DISPLAY As String = "HYE Support"           ' 패널 표시 이름
        Private Const PANEL_INTERNAL As String = "id_hyeSupportPanel"   ' 패널 InternalName

        ' 현재 떠 있는 폼 인스턴스 (여러 번 생성 방지)
        Private _form As CreateSupportForm


#Region "ApplicationAddInServer Members"

        Public Sub Activate(addInSiteObject As ApplicationAddInSite,
              firstTime As Boolean) Implements ApplicationAddInServer.Activate
            Try
                ' Inventor Application 핸들 저장
                Globals.g_inventorApplication = addInSiteObject.Application

                ' 리본 리셋 이벤트 구독
                m_uiEvents = Globals.g_inventorApplication.
              UserInterfaceManager.UserInterfaceEvents

                Dim cmdMgr As CommandManager =
                  Globals.g_inventorApplication.CommandManager
                Dim controlDefs As ControlDefinitions = cmdMgr.ControlDefinitions

                ' ─ 버튼 정의 재사용 or 생성 ─
                Try
                    ' 이미 정의된 버튼이 있으면 그걸 재사용
                    m_supportButton = CType(controlDefs.Item("hye.CreateSupportButton.v2"), ButtonDefinition)

                Catch
                    ' ▼▼ 여기서 새로 버튼을 만들 때 Resource1.surpport32 아이콘을 함께 지정 ▼▼

                    ' 1) 리소스에서 32x32 PNG 가져오기
                    Dim baseBmp As Drawing.Bitmap = Nothing
                    Try
                        ' Resource1.resx 안의 surpport32.png 사용
                        baseBmp = CreateSupport.Resource1.support32
                    Catch
                        baseBmp = Nothing
                    End Try

                    ' 2) 16x16 / 32x32 비트맵 만들기 (리본 작은 아이콘 / 큰 아이콘용)
                    Dim smallIconDisp As stdole.IPictureDisp = Nothing
                    Dim largeIconDisp As stdole.IPictureDisp = Nothing

                    If baseBmp IsNot Nothing Then
                        Try
                            ' 작은 아이콘(16x16)
                            Dim smallBmp As New Drawing.Bitmap(baseBmp, New Drawing.Size(16, 16))
                            ' 큰 아이콘(32x32)
                            Dim largeBmp As New Drawing.Bitmap(baseBmp, New Drawing.Size(32, 32))

                            ' PictureDispConverter 를 이용해 IPictureDisp 로 변환
                            ' ※ PictureDispConverter.vb 안의 Shared 함수 이름이
                            '    ImageToPictureDisp 인 기준으로 작성
                            smallIconDisp = PictureDispConverter.ImageToPictureDisp(smallBmp)
                            largeIconDisp = PictureDispConverter.ImageToPictureDisp(largeBmp)

                        Catch
                            ' 아이콘 변환에 실패하면 아이콘 없이 생성하도록 둠
                            smallIconDisp = Nothing
                            largeIconDisp = Nothing
                        End Try
                    End If

                    ' 3) 아이콘(있으면) 포함해서 버튼 정의 생성
                    m_supportButton = controlDefs.AddButtonDefinition(
                        DisplayName:="Create Support",                         ' 리본에 보일 글자
                        InternalName:="hye.CreateSupportButton.v2",               ' 고유 이름(ID)
                        Classification:=CommandTypesEnum.kShapeEditCmdType,    ' 명령 종류(도형 편집)
                        ClientId:=Globals.AddInClientID(),                     ' Add-In GUID
                        DescriptionText:="서포트 자동 생성 도구",              ' 설명(상태줄 등)
                        ToolTipText:="Create Support",                         ' 툴팁
                        StandardIcon:=smallIconDisp,                           ' 리본 작은 아이콘
                        LargeIcon:=largeIconDisp)                              ' 리본 큰 아이콘
                End Try



                ' 버튼 클릭 이벤트 연결
                AddHandler m_supportButton.OnExecute,
             AddressOf m_supportButton_OnExecute

                AddToUserInterface()

            Catch ex As Exception
                ' 필요하면 로그 추가
            End Try
        End Sub

        Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
            Try
                ' 1) 버튼 클릭 이벤트 핸들러 해제
                '    - Activate에서 AddHandler로 연결했으므로
                '      Deactivate에서 RemoveHandler로 정리해 주면 메모리 누수 방지에 도움이 된다.
                If m_supportButton IsNot Nothing Then
                    Try
                        RemoveHandler m_supportButton.OnExecute, AddressOf m_supportButton_OnExecute
                    Catch
                        ' 버튼이 이미 해제된 상태일 수도 있으므로 조용히 무시
                    End Try
                    m_supportButton = Nothing
                End If

                ' 2) UI 이벤트 객체 해제
                '    - WithEvents 참조를 끊어주면 리본 이벤트 구독도 정리된다.
                m_uiEvents = Nothing

                ' 3) 전역 Application 참조 해제
                '    - Inventor 종료 시점을 깔끔하게 맞추기 위해 Nothing으로 설정
                Globals.g_inventorApplication = Nothing

                ' 4) 폼 참조 정리
                '    - 혹시라도 폼이 떠 있는 상태에서 Deactivate가 호출되면
                '      참조를 끊어 GC가 수거할 수 있게 한다.
                If _form IsNot Nothing Then
                    Try
                        ' 필요하면 여기서 _form.Close()를 호출해도 된다.
                        ' 지금은 단순히 참조만 제거.
                        RemoveHandler _form.FormClosed, AddressOf FormClosed_Handler
                    Catch
                    End Try
                    _form = Nothing
                End If

            Catch
                ' Deactivate에서 예외가 터져도 Inventor 종료를 막지 않도록 조용히 처리
            End Try

            ' 5) 가비지 컬렉션
            '    - Add-In 언로드 직후 한 번 강제 호출해서
            '      관리되는 리소스를 빠르게 정리하도록 유도
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Sub


        Public ReadOnly Property Automation As Object _
                Implements ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        Public Sub ExecuteCommand(commandID As Integer) _
                Implements ApplicationAddInServer.ExecuteCommand
            ' 사용 안 함
        End Sub

#End Region


#Region "리본 UI 구성"

        Private Sub AddToUserInterface()
            Dim uim As UserInterfaceManager =
              Globals.g_inventorApplication.UserInterfaceManager

            ' Part 리본
            Try
                Dim partRibbon As Ribbon = uim.Ribbons.Item("Part")
                Dim hyTab As RibbonTab =
                  GetOrCreateTab(partRibbon, TAB_DISPLAY, TAB_INTERNAL)

                Dim supportPanel As RibbonPanel =
                  GetOrCreatePanel(hyTab, PANEL_DISPLAY, PANEL_INTERNAL)

                AddOrGetButton(supportPanel, m_supportButton)
            Catch
            End Try

            ' Assembly 리본
            Try
                Dim asmRibbon As Ribbon = uim.Ribbons.Item("Assembly")
                Dim hyTab As RibbonTab =
                  GetOrCreateTab(asmRibbon, TAB_DISPLAY, TAB_INTERNAL)

                Dim supportPanel As RibbonPanel =
                  GetOrCreatePanel(hyTab, PANEL_DISPLAY, PANEL_INTERNAL)

                AddOrGetButton(supportPanel, m_supportButton)
            Catch
            End Try
        End Sub

        Private Function GetOrCreateTab(ribbon As Ribbon,
                        display As String,
                        internalName As String) As RibbonTab

            For Each t As RibbonTab In ribbon.RibbonTabs
                If t.InternalName = internalName Then
                    Return t
                End If
            Next

            Return ribbon.RibbonTabs.Add(display, internalName, Globals.AddInClientID())
        End Function

        Private Function GetOrCreatePanel(tab As RibbonTab,
                         display As String,
                         internalName As String) As RibbonPanel

            For Each p As RibbonPanel In tab.RibbonPanels
                If p.InternalName = internalName Then
                    Return p
                End If
            Next

            Return tab.RibbonPanels.Add(display, internalName, Globals.AddInClientID())
        End Function

        Private Sub AddOrGetButton(panel As RibbonPanel,
                           def As ButtonDefinition)

            ' ★ 방어 코드: 버튼 정의가 없으면(생성 실패 등) 그냥 아무 것도 하지 않고 종료
            '   - NullReferenceException 방지용
            If def Is Nothing Then
                Exit Sub
            End If

            ' 이미 같은 InternalName의 버튼이 있으면 다시 추가하지 않음
            For Each c As CommandControl In panel.CommandControls
                If c.InternalName = def.InternalName Then
                    Exit Sub
                End If
            Next

            ' 패널에 버튼 추가
            panel.CommandControls.AddButton(def, True, True)
        End Sub


        ' 리본 리셋 시 다시 패널/버튼 추가
        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) _
            Handles m_uiEvents.OnResetRibbonInterface

            AddToUserInterface()
        End Sub

#End Region


#Region "버튼 클릭 → CreateSupportForm 실행"

        ' ============================================================
        ' 버튼 클릭 → (임시) 메시지 박스만 띄워보기
        '   - 폼 생성 / Owner 래핑 / Inventor 핸들 사용 전부 제거
        '   - 이 상태에서 여전히 인벤터가 죽으면, 문제는 Add-In이 아니라
        '     Inventor나 다른 환경 쪽에 더 가깝다.
        ' ============================================================
        Private Sub m_supportButton_OnExecute(Context As NameValueMap)
            Try
                ' ① Inventor Application 확인
                If Globals.g_inventorApplication Is Nothing Then
                    WinForms.MessageBox.Show(
                "Inventor Application이 유효하지 않습니다." & vbCrLf &
                "Add-In이 올바르게 로드되었는지 확인하세요.",
                "CreateSupport",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
                    Return
                End If

                ' ② Inventor 메인 윈도우를 Owner로 래핑
                Dim mainHwnd As IntPtr = New IntPtr(Globals.g_inventorApplication.MainFrameHWND)
                Dim owner As New Globals.WindowWrapper(mainHwnd)

                ' ③ 이미 폼이 떠 있으면 재활성화
                If _form IsNot Nothing AndAlso Not _form.IsDisposed Then
                    If _form.WindowState = WinForms.FormWindowState.Minimized Then
                        _form.WindowState = WinForms.FormWindowState.Normal
                    End If
                    _form.Activate()
                    _form.BringToFront()
                    Return
                End If

                ' ④ 새 폼 생성
                _form = New CreateSupportForm()
                AddHandler _form.FormClosed, AddressOf FormClosed_Handler

                ' Owner(=Inventor 창) 기준 중앙
                _form.StartPosition = FormStartPosition.CenterParent

                ' ⑤ Owner를 지정해서 표시 → Inventor가 포커스를 가져가도
                '    Owner 위에 계속 떠 있는 모델리스 창처럼 동작
                _form.Show(owner)

            Catch ex As Exception
                WinForms.MessageBox.Show(
            "폼을 여는 중 오류가 발생했습니다." & vbCrLf &
            ex.Message,
            "CreateSupport - OnExecute 오류",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)
            End Try
        End Sub





        Private Sub FormClosed_Handler(sender As Object,
                       e As FormClosedEventArgs)
            _form = Nothing
        End Sub

#End Region

    End Class
End Namespace


' ============================================================
' Globals 모듈
'  - g_inventorApplication : 전체 Add-In에서 공유하는 Application
'  - AddInClientID()       : 위 GuidAttribute에서 GUID를 읽어옴
'  - WindowWrapper         : Inventor 창을 Owner로 쓰기 위한 래퍼
' ============================================================
Public Module Globals

    Public g_inventorApplication As Inventor.Application

    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(CreateSupport.StandardAddInServer)
            Dim attrs() As Object =
              t.GetCustomAttributes(GetType(GuidAttribute), False)

            Dim guidAttr As GuidAttribute =
              CType(attrs(0), GuidAttribute)

            guid = "{" & guidAttr.Value.ToString() & "}"
        Catch
        End Try
        Return guid
    End Function

    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        Private _hwnd As IntPtr

        Public Sub New(handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle As IntPtr _
                Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property
    End Class

End Module