Attribute VB_Name = "Global"
Global TV As TVEngine
Global Scene As TVScene
Global TextureFactory As TVTextureFactory
Global InputEngine As TVInputEngine
Global Atmosphere As TVAtmosphere
Global Tank As TVActor2
Global Player As TVActor2
Global Screen2DImmediate As TVScreen2DImmediate
Global Screen2DText As TVScreen2DText
Global CollisionResult As TVCollisionResult

Global SoundEngine As New TV3DMedia.TVSoundEngine
Global Listener As TVListener
Global SFX(0 To 25) As New TV3DMedia.TVSoundWave3D

Global Bomb(9) As TVMesh
Global BombFrame(9) As Integer
Global BombIsFree(9) As Boolean
Global BombFree As Integer

Global tmpMouseX As Long, tmpMouseY As Long
Global tmpMouseB1 As Integer, tmpMouseB2 As Integer, tmpMouseB3 As Integer
Global tmpMouseScrollOld As Long, tmpMouseScrollNew As Long
Global sngAngleX As Single
Global sngAngleY As Single
Global sngPositionX As Single
Global sngPositionY As Single
Global sngPositionZ As Single
Global snglookatX As Single
Global snglookatY As Single
Global snglookatZ As Single
Global HealthPlayer As Integer
Global WeaponPlayer As Integer

Sub LoadSkyBox(ID As String)
DoEvents
Scene.LoadTexture App.Path + "\Data\Pic\" + ID + "bk.jpg", , , "bk"
Scene.LoadTexture App.Path + "\Data\Pic\" + ID + "dn.jpg", , , "dn"
Scene.LoadTexture App.Path + "\Data\Pic\" + ID + "ft.jpg", , , "ft"
Scene.LoadTexture App.Path + "\Data\Pic\" + ID + "lf.jpg", , , "lf"
Scene.LoadTexture App.Path + "\Data\Pic\" + ID + "up.jpg", , , "up"
Scene.LoadTexture App.Path + "\Data\Pic\" + ID + "rt.jpg", , , "rt"
Atmosphere.SkyBox_SetTexture GetTex("ft"), GetTex("bk"), GetTex("lf"), GetTex("rt"), GetTex("up"), GetTex("dn")
Atmosphere.SkyBox_Enable True
DoEvents
End Sub

Sub CameraRotation()
DoEvents
tmpMouseScrollOld = tmpMouseScrollNew
InputEngine.GetMouseState tmpMouseX, tmpMouseY, tmpMouseB1, tmpMouseB2, tmpMouseB3, tmpMouseScrollNew
sngAngleX = sngAngleX - (tmpMouseY / 100)
sngAngleY = sngAngleY - (tmpMouseX / 100)
If sngAngleX > 1.3 Then sngAngleX = 1.3
If sngAngleX < -1.3 Then sngAngleX = -1.3
snglookatX = sngPositionX + Cos(sngAngleY)
snglookatY = sngPositionY + Tan(sngAngleX)
snglookatZ = sngPositionZ + Sin(sngAngleY)
Player.SetRotation 0, 180 - Rad2Deg(sngAngleY), -Rad2Deg(sngAngleX)
Scene.SetCamera sngPositionX, sngPositionY, sngPositionZ, snglookatX, snglookatY, snglookatZ
DoEvents
End Sub

Sub Close3D()
DoEvents
Set TV = Nothing
Set Scene = Nothing
Set InputEngine = Nothing
Set Atmosphere = Nothing
Set Tank = Nothing
Set Screen2DImmediate = Nothing
Set TextureFactory = Nothing
Set Player = Nothing
Set CollisionResult = Nothing
DoEvents
End Sub

Sub IntBomb()
Dim i As Integer
For i = 1 To 8
TextureFactory.LoadTexture App.Path + "\Data\Pic\expl" + Format(i) + ".jpg", "expl" + Format(i), , , TV_COLORKEY_BLACK
Next
For i = 0 To 9
Set Bomb(i) = Scene.CreateMeshBuilder
Bomb(i).AddWall GetTex("expl1"), -1, 0, 1, 0, 2, -1
Bomb(i).AddFloor GetTex("expl1"), -1, -1, 1, 1
BombFrame(i) = 0
BombIsFree(i) = True
Bomb(i).SetPosition -10000, -10000, -10000
Next
End Sub

Sub NewBomb(Pos As D3DVECTOR, Size As Single)
BombIsFree(BombFree) = False
Bomb(BombFree).SetPosition Pos.x, Pos.Y, Pos.Z
Bomb(BombFree).ScaleMesh Size, Size, Size
Dim i As Integer
For i = 0 To 9
If BombIsFree(i) = True Then
BombFree = i
Else
BombFree = 0
End If
Next
End Sub

Sub UpdateBomb()
Dim i As Integer
For i = 0 To 9
If BombIsFree(i) = False Then
If BombFrame(i) = 8 Then
BombFrame(i) = 0
BombIsFree(i) = True
Bomb(i).SetPosition -10000, -10000, -10000
Else
BombFrame(i) = BombFrame(i) + 1
Bomb(i).SetTexture GetTex("expl" + Format(BombFrame(i)))
Bomb(i).LookAtPoint Scene.GetCamera.GetPosition
Bomb(i).RotateX 90
End If
End If
Next
End Sub

Sub IntSound()
SoundEngine.Init GM.hWnd
Set Listener = SoundEngine.Get3DListener
Dim i As Integer
For i = 0 To 25
SFX(i).Load App.Path + "\Data\SOUND\" + Format(i) + ".wav"
SFX(i).Loop_ = False
Next
End Sub
