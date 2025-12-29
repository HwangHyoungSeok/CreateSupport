' ============================================
' PictureDispConverter.vb
'  - System.Drawing.Image → stdole.IPictureDisp 변환용
'  - Inventor 리본 버튼 아이콘에 쓰기 위한 도우미 클래스
' ============================================

Imports System.Drawing                     ' Bitmap, Image 사용
Imports System.Runtime.InteropServices      ' DllImport, Marshal
Imports stdole                              ' IPictureDisp 인터페이스

Public Class PictureDispConverter

    ' --------------------------------------------
    ' OLE 내부 API 선언 (이미지를 IPictureDisp 로 감싸줌)
    '  - OleCreatePictureIndirect 는 Win32 함수라 DllImport 필요
    ' --------------------------------------------
    <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect",
               SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Shared Function OleCreatePictureIndirect(
    <[In]> ByVal picDesc As PICTDESC,           ' 그림 설명 구조체
    <[In]> ByRef iid As System.Guid,            ' ★ System.Guid 로 명시
    <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean
) As IPictureDisp
    End Function

    ' --------------------------------------------
    ' IPictureDisp 를 만들기 위한 구조체 정의
    '  - PICTDESC 는 OLE 쪽에서 요구하는 그림 정보 형식
    ' --------------------------------------------
    <StructLayout(LayoutKind.Sequential)>
    Private Class PICTDESC

        ' 구조체 크기 (Marshal.SizeOf 로 자동 계산)
        Friend cbSizeofstruct As Integer = Marshal.SizeOf(GetType(PICTDESC))

        ' 그림 타입 (1 = Bitmap)
        Friend picType As Integer

        ' 실제 그림 핸들 (HBITMAP)
        Friend hImage As IntPtr

        ' 여백용 필드 (사용 안 함)
        Friend x As Integer
        Friend y As Integer

        Private Sub New()
            ' 외부에서 new 금지 → 팩토리 메서드 사용
        End Sub

        ' Bitmap 으로부터 PICTDESC 생성
        Public Shared Function CreateBitmapPictureDesc(bmp As Bitmap) As PICTDESC
            Dim desc As New PICTDESC()
            desc.picType = 1                 ' 1 = PICTYPE_BITMAP
            desc.hImage = bmp.GetHbitmap()   ' GDI 비트맵 핸들 얻기
            Return desc
        End Function
    End Class

    ' IPictureDisp 인터페이스의 GUID (OleCreatePictureIndirect 에 전달)
    Private Shared ReadOnly iPictureDispGuid As System.Guid =
    GetType(stdole.IPictureDisp).GUID   ' ★ System.Guid 로 명시

    ' --------------------------------------------------------
    ' Public Shared Function ImageToPictureDisp
    '  - WinForms 의 Image/Bitmap → 리본 아이콘용 IPictureDisp 변환
    '  - 예: PictureDispConverter.ImageToPictureDisp(myBitmap)
    ' --------------------------------------------------------
    Public Shared Function ImageToPictureDisp(image As Image) As IPictureDisp
        ' 1) NULL 체크
        If image Is Nothing Then
            Return Nothing
        End If

        ' 2) Image 를 Bitmap 으로 캐스팅
        Dim bmp As Bitmap = TryCast(image, Bitmap)
        If bmp Is Nothing Then
            ' Image 가 Bitmap 이 아니면 새 Bitmap 으로 한 번 감싸줌
            bmp = New Bitmap(image)
        End If

        ' 3) Bitmap → PICTDESC 구조체 생성
        Dim pictDesc As PICTDESC = PICTDESC.CreateBitmapPictureDesc(bmp)

        ' 4) OLE 함수 호출해서 IPictureDisp 생성
        Dim pict As IPictureDisp = OleCreatePictureIndirect(pictDesc, iPictureDispGuid, True)

        ' 5) IPictureDisp 반환 (리본 StandardIcon / LargeIcon 에 사용)
        Return pict
    End Function

End Class
