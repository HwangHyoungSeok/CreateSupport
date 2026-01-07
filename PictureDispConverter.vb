' PictureDispConverter.vb
Imports System.Runtime.InteropServices
Imports stdole
Imports Drawing = System.Drawing

Public NotInheritable Class PictureDispConverter
    Private Sub New()
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure PICTDESC
        Public cbSizeOfStruct As Integer
        Public picType As Integer
        Public hImage As IntPtr
        Public xExt As Integer
        Public yExt As Integer
    End Structure

    <DllImport("oleaut32.dll", PreserveSig:=False)>
    Private Shared Function OleCreatePictureIndirect(
    ByRef pictdesc As PICTDESC,
    ByRef iid As System.Guid,
    ByVal fOwn As Boolean) As IPictureDisp
    End Function


    Public Shared Function ImageToPictureDisp(img As Drawing.Image) As IPictureDisp
        If img Is Nothing Then Return Nothing

        Dim bmp As Drawing.Bitmap = TryCast(img, Drawing.Bitmap)
        If bmp Is Nothing Then
            bmp = New Drawing.Bitmap(img)
        End If

        Dim hBmp As IntPtr = bmp.GetHbitmap() ' GDI 핸들
        Dim pd As New PICTDESC With {
            .cbSizeOfStruct = Marshal.SizeOf(GetType(PICTDESC)),
            .picType = 1, ' PICTYPE_BITMAP = 1
            .hImage = hBmp,
            .xExt = 0,
            .yExt = 0
        }

        Dim ipic As IPictureDisp = Nothing
        Dim iid As System.Guid = GetType(stdole.IPictureDisp).GUID
        ipic = OleCreatePictureIndirect(pd, iid, True)
        Return ipic
    End Function
End Class
