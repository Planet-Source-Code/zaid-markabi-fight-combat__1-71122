VERSION 5.00
Begin VB.Form GM 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "GM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Zoom As Boolean
Dim TankGo As Boolean
Dim Tankwd As Boolean
Dim Tankwd2 As Single
Dim TankPosition As D3DVECTOR
Dim TankDestination As D3DVECTOR
Dim TankDirection As D3DVECTOR
Dim Reloading As Integer
Dim OhWait As Integer
Dim PlaneHealth As Integer
Dim EndGame As Boolean
Dim PlaneSpeedFly As Integer

Private Sub Form_Load()
Me.Hide
Set TV = New TVEngine
TV.SetDebugMode False
TV.SetAngleSystem TV_ANGLE_DEGREE
DoEvents
Set Scene = New TVScene
Set InputEngine = New TVInputEngine
Set Atmosphere = New TVAtmosphere
Set Tank = New TVActor2
Set Player = New TVActor2
Set Screen2DImmediate = New TVScreen2DImmediate
Set TextureFactory = New TVTextureFactory
Set CollisionResult = New TVCollisionResult
Set Screen2DText = New TVScreen2DText

TV.Init3DWindowedMode Picture1.hWnd
LoadSkyBox ""
Scene.SetViewFrustum 90, 10000

TextureFactory.LoadTexture App.Path + "\Data\Pic\PLANE.bmp", "PLANE", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Pic\crossh.bmp", "crossh", , , TV_COLORKEY_BLACK
Scene.LoadTexture App.Path + "\Data\Pic\HEALTH.bmp", , , "HEALTH"
TextureFactory.LoadTexture App.Path + "\Data\Pic\BULLET.bmp", "BULLET", , , TV_COLORKEY_BLACK
TextureFactory.LoadTexture App.Path + "\Data\Pic\NSWE.bmp", "NSWE", , , TV_COLORKEY_GREEN

Tank.Load App.Path + "\Data\Model\apache.mdl"
Tank.SetScale 0.1, 0.1, 0.1
Tank.SetAnimationID 0
Tank.SetAnimationLoop True
sngPositionX = Rnd * 200
sngPositionY = Rnd * 150
sngPositionZ = Rnd * 200
Tank.PlayAnimation 10

Player.Load App.Path + "\Data\Model\v_ak47.mdl"
Player.SetScale 0.3, 0.3, 0.3
Player.SetAnimationID 0
Player.SetAnimationLoop True
Player.SetPosition sngPositionX, sngPositionY, sngPositionZ
Player.PlayAnimation 10

Reloading = 2

PlaneSpeedFly = 1

HealthPlayer = 100
WeaponPlayer = 10
PlaneHealth = 9

IntBomb

IntSound

SFX(0).Loop_ = True
SFX(0).Play

Screen2DText.NormalFont_Create "I", "Arial", 16, True, False, False

Me.Show
Do
DoEvents

If InputEngine.IsKeyPressed(TV_KEY_ESCAPE) = True Then Exit Do

CameraRotation

MoveTank

Fire

UpdateBomb

TV.Clear
Atmosphere.Atmosphere_Render
Tank.Render
Scene.RenderAllMeshes

Screen2DImmediate.ACTION_Begin2D
Screen2DImmediate.DRAW_Texture GetTex("PLANE"), 0, 0, Me.ScaleWidth, Me.ScaleHeight
If Zoom = False Then
Screen2DImmediate.DRAW_TextureRotated GetTex("crossh"), (Me.ScaleWidth / 2), (Me.ScaleHeight / 2), 30, 30, Rad2Deg(sngAngleY)
Else
Screen2DImmediate.DRAW_TextureRotated GetTex("crossh"), (Me.ScaleWidth / 2), (Me.ScaleHeight / 2), 40, 40, Rad2Deg(sngAngleY), , , , , RGBA(1, 0, 0, 1), RGBA(1, 0, 0, 1), RGBA(1, 0, 0, 1), RGBA(1, 0, 0, 1)
End If
Screen2DImmediate.DRAW_TextureRotated GetTex("NSWE"), 80, 80, 150, 150, Rad2Deg(sngAngleY)
Screen2DImmediate.DRAW_Texture GetTex("HEALTH"), 25, Me.ScaleHeight - 100, 100, Me.ScaleHeight - 25
Screen2DImmediate.DRAW_Texture GetTex("BULLET"), 25, Me.ScaleHeight - 200, 100, Me.ScaleHeight - 125
Screen2DImmediate.ACTION_End2D

Player.Render

Screen2DText.ACTION_BeginText
Screen2DText.NormalFont_DrawText Format(HealthPlayer) + " %", 110, Me.ScaleHeight - 70, RGBA(1, 0, 0, 1), "I"
Screen2DText.NormalFont_DrawText Format(WeaponPlayer), 110, Me.ScaleHeight - 170, RGBA(1, 0, 0, 1), "I"
Screen2DText.ACTION_EndText

TV.RenderToScreen
Loop

Close3D
Unload Me
End Sub

Private Sub Form_Resize()
Picture1.Width = Me.ScaleWidth
Picture1.Height = Me.ScaleHeight + 40
TV.ResizeDevice
End Sub

Private Sub MoveTank()
If TankGo = False Then
TankDestination = Vector(Rnd * 400, Rnd * 200, Rnd * 400)
TankGo = True
Else

If GetDistance3D(TankPosition.x, TankPosition.Y, TankPosition.Z, TankDestination.x, TankDestination.Y, TankDestination.Z) > 4 Then
TankPosition = VAdd(Tank.GetPosition, VScale(TankDirection, PlaneSpeedFly))
TankDirection = VNormalize(VSubtract(TankDestination, Tank.GetPosition))
Tank.SetPosition TankPosition.x, TankPosition.Y, TankPosition.Z
Tank.Lookat sngPositionX, sngPositionY, sngPositionZ
Tank.RotateY 90 + 0.01 * TV.TickCount
If Tankwd = False Then
Tank.RotateX Tankwd2
Tankwd2 = Tankwd2 + Rnd
Else
Tank.RotateX Tankwd2
Tankwd2 = Tankwd2 - Rnd
End If
If Tankwd2 > 40 Or Tankwd2 < -80 Then
If Tankwd = False Then
Tankwd = True
Else
Tankwd = False
End If
End If
Else
If EndGame = False Then
TankGo = False
Else
Win.Show
Close3D
Unload Me
End If
End If

End If

If GetDistance3D(TankPosition.x, TankPosition.Y, TankPosition.Z, sngPositionX, sngPositionY, sngPositionZ) < 50 Then
If OhWait > 0 Then
OhWait = OhWait - 1
Else
If Rnd > 0.5 Then
SFX(5).Play
Else
SFX(6).Play
End If
OhWait = 50
HealthPlayer = HealthPlayer - Int(Rnd * 50)
End If
End If
End Sub

Private Sub Fire()
If Reloading = 2 Then
Set CollisionResult = Scene.MousePicking(Me.ScaleWidth / 2, Me.ScaleHeight / 2, TV_COLLIDE_ACTOR2, TV_TESTTYPE_ACCURATETESTING)
Zoom = CollisionResult.IsCollision
If tmpMouseB1 <> 0 Then
SFX(2).Play
Player.SetAnimationID 4
Player.SetAnimationLoop False
Player.PlayAnimation 100
Reloading = 0
WeaponPlayer = WeaponPlayer - 1
If Zoom = True Then
CollisionResult.GetCollisionActor2.RotateY 10
CollisionResult.GetCollisionActor2.RotateX 10
NewBomb CollisionResult.GetCollisionActor2.GetPosition, 25
TankGo = False
If Rnd > 0.5 Then
SFX(1).Play
Else
SFX(4).Play
End If
PlaneHealth = PlaneHealth - 1
End If
End If
End If

If Reloading = 0 And Player.IsAnimationFinished = True Then
Reloading = 1
SFX(3).Play
Player.SetAnimationID 1
Player.SetAnimationLoop False
Player.PlayAnimation 75
End If

If Reloading = 1 And Player.IsAnimationFinished = True Then
Reloading = 2
Player.SetAnimationID 0
Player.SetAnimationLoop False
Player.PlayAnimation 10
End If

If PlaneHealth < 1 Then
NewBomb TankPosition, 25
If EndGame = False Then
Tank.Load App.Path + "\Data\Model\dead_osprey.mdl"
Tank.SetScale 0.1, 0.1, 0.1
Tank.SetPosition TankPosition.x, TankPosition.Y, TankPosition.Z
TankDestination = Vector(Rnd * 400, -500, Rnd * 400)
TankGo = True
Tank.SetAnimationID 0
Tank.SetAnimationLoop True
Tank.PlayAnimation 10
EndGame = True
PlaneSpeedFly = PlaneSpeedFly * 2
End If
End If

If WeaponPlayer < 1 And PlaneHealth > 0 Then
Lose.Show
Close3D
Unload Me
End If
End Sub
