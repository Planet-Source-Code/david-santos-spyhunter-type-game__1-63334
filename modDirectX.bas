Attribute VB_Name = "modDirectX"
Public DX7 As New DirectX7
Public DSound As DirectSound
Public DDraw As DirectDraw7
Public DDClip As DirectDrawClipper

Public ddsPrimary As DirectDrawSurface7
Public ddsBuffer As DirectDrawSurface7
Public ddsBG As DirectDrawSurface7

Public ddsRoad As DirectDrawSurface7
Public ddsObjects As DirectDrawSurface7
Public ddsOther As DirectDrawSurface7

Public ddsd As DDSURFACEDESC2

Public Sub InitDX(myForm As Form)
    Set DDraw = DX7.DirectDrawCreate("")
    ' Set DSound = DX7.DirectSoundCreate("")
    DDraw.SetCooperativeLevel myForm.hWnd, DDSCL_NORMAL
        
    Dim hwCaps As DDCAPS
    Dim helCaps As DDCAPS
    DDraw.GetCaps hwCaps, helCaps
  
    If (hwCaps.lFXCaps And DDFXCAPS_BLTROTATION) = 0 Then
        Debug.Print "HW: Rotation not supported"
    End If
    
    If (helCaps.lFXCaps And DDFXCAPS_BLTROTATION) = 0 Then
        Debug.Print "HEL: Rotation not supported"
    End If
    
End Sub

Public Sub DestroyDX()
 Set DDraw = Nothing
 Set DSound = Nothing
End Sub
