VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3240
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4995
   Icon            =   "mY3D2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   " Day"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   8280
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   " Night"
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   8280
      Value           =   -1  'True
      Width           =   855
   End
   Begin RMControl7.RMCanvas RMCanvas1 
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5530
   End
   Begin VB.Menu filemenu 
      Caption         =   "File"
      Begin VB.Menu bigmenu 
         Caption         =   "Big WIndow"
      End
      Begin VB.Menu smallmenu 
         Caption         =   "Small Window"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Direct 3D RM Game v 1.0 by: Radeon - my first direct3d game!!!
'This looks like a lot of code, but dont be frightened
'it is easy to follow
'******************************************************
'This game runs well on a 100mhz PC in the small window,
'if you have something higher then 100mhz I recommend you
'run it in the big window, the game will look much better
'Go to "File", "Big Window" to run in full screen
'******************************************************
'Direct 3D RM Game v 1.0 Copyright (c) 2001 Radeon Melnyk
'If you use any of this code in any of your programs, please mention
'my name there
'I must mention DAVE for his cool 3D Park game-that's where
'I got my .X objects, and learned how to load them in
'I also must mention www.MSDN.microsoft.com/directx
'- DIRECTX RULES!!!!
'**************************************************************
'Whatch out for Version 2.0, some of the things I hope to
'include are, some more buildings and houses, more roads,
'Collision Detection, *Sreet Lights*, Moving sprites, and guns.
'I AM ALSO WORKING ON A 3D RACING GAME, BUT THAT WON'T BE OUT
'FOR A WHILE.
'**************************************************************

Option Explicit
Public DX As New DirectX7
Dim Surface As Direct3DRMMeshBuilder3
Dim Surfacefrm As Direct3DRMFrame3
Dim Lspot As Direct3DRMLight
Dim Lspotfrm As Direct3DRMFrame3
Dim Lspotm As Direct3DRMMeshBuilder3
Dim Amblight As Direct3DRMLight
Dim Ambm As Direct3DRMMeshBuilder3
Dim Ambfrm As Direct3DRMFrame3
'Dim Lspotmfrm(0 To 10) As Direct3DRMFrame3
'the walls, which make up the house
Dim Walla As Direct3DRMMeshBuilder3
Dim Wallb As Direct3DRMMeshBuilder3
Dim Wallc As Direct3DRMMeshBuilder3
Dim Walld As Direct3DRMMeshBuilder3
Dim Walle As Direct3DRMMeshBuilder3
'wall frames of the house
Dim Wallfrma As Direct3DRMFrame3
Dim Wallfrmb As Direct3DRMFrame3
Dim Wallfrmc As Direct3DRMFrame3
Dim Wallfrmd As Direct3DRMFrame3
Dim Wallfrme As Direct3DRMFrame3
'The walls of the Building
Dim bWall1 As Direct3DRMMeshBuilder3
Dim bWall2 As Direct3DRMMeshBuilder3
Dim bWall3 As Direct3DRMMeshBuilder3
Dim bWall4 As Direct3DRMMeshBuilder3
Dim bRoof As Direct3DRMMeshBuilder3
'Frames of the Building
Dim bWallfrm1 As Direct3DRMFrame3
Dim bWallfrm2 As Direct3DRMFrame3
Dim bWallfrm3 As Direct3DRMFrame3
Dim bWallfrm4 As Direct3DRMFrame3
Dim bRooffrm As Direct3DRMFrame3
'parking lot
Dim Lot As Direct3DRMMeshBuilder3
Dim Lotfrm As Direct3DRMFrame3
'the bench
Dim BenchMesh As Direct3DRMMeshBuilder3
Dim Benchfrm As Direct3DRMFrame3
'light post
'Dim lp As Direct3DRMMeshBuilder3
'Dim lpfrm As Direct3DRMFrame3
'car
'Dim carmesh As Direct3DRMMeshBuilder3
'Dim carfrm As Direct3DRMFrame3
'the trees
Dim TreeTopMesh As Direct3DRMMeshBuilder3
Dim TreeBotMesh As Direct3DRMMeshBuilder3
Dim TreeTopfrm As Direct3DRMFrame3
Dim TreeBotfrm As Direct3DRMFrame3
'look up or down
Dim UpDo As Single
Const UpDoSpeed = 0.05
Dim matUpDo As D3DMATRIX
'for textures
Dim Stexture As Direct3DRMTexture3
Dim texture As Direct3DRMTexture3
Dim Btexture As Direct3DRMTexture3
Dim Btexture2 As Direct3DRMTexture3
Dim Ltexture As Direct3DRMTexture3
'Road stuff
Dim Road As Direct3DRMMeshBuilder3
Dim RoadFrame As Direct3DRMFrame3
'lights
Dim Lights As Direct3DRMMeshBuilder3
'Dim CameraBoxFrame As Direct3DRMFrame3
Dim light As Direct3DRMLight
Dim Lightsfrm(0 To 10) As Direct3DRMFrame3
Dim Lightfrm As Direct3DRMFrame3
Dim Zaxis As Integer
Dim V As D3DVECTOR
Dim go
Const PI = 3.1415926535 'pi
Private Sub bigmenu_Click()
Form2.Height = Screen.Height
Form2.Width = Screen.Width
Form2.Top = Screen.TwipsPerPixelY
Form2.Left = Screen.TwipsPerPixelX
RMCanvas1.Width = Form2.Width
RMCanvas1.Height = Form2.Height
RMCanvas1.Left = 0
End Sub
Private Sub smallmenu_Click()
Form2.Width = 5115
Form2.Height = 5040
RMCanvas1.Width = Form1.Width
RMCanvas1.Height = 3135
RMCanvas1.Left = 0
End Sub
Private Sub Form_Load()
   'turn = 0
    'MsgBox "USE THE ARROW KEYS TO MOVE"
    RMCanvas1.StartWindowed
    Show
    Zaxis = 0
If Form1.Option1.Value = True Then Form2.Option1.Value = True
If Form1.Option2.Value = True Then Form2.Option2.Value = True
    
    RMCanvas1.Viewport.SetBack 10000 'how far you can see
   If Option2.Value = True Then RMCanvas1.SceneFrame.SetSceneBackgroundRGB 0, 0, 0.3 'background color
   If Option1.Value = True Then RMCanvas1.SceneFrame.SetSceneBackgroundRGB 0, 1, 1 'background color
'********************************************************************************************************
'PRIVATE SUB TREES_EVERYTHING() 'dont mind this, its just the way I organize my code
    'create tree meshes
  Set TreeTopMesh = RMCanvas1.D3DRM.CreateMeshBuilder()
  Set TreeBotMesh = RMCanvas1.D3DRM.CreateMeshBuilder()
  Set TreeTopfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
  Set TreeBotfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
  'load trees
    TreeTopMesh.LoadFromFile "treetop.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    TreeBotMesh.LoadFromFile "treebot.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    'set tree positions
    TreeTopfrm.SetPosition Nothing, 20, -62, -27
    TreeBotfrm.SetPosition Nothing, 20, -60.2, -27
    'set material
    TreeTopfrm.SetMaterialMode D3DRMMATERIAL_FROMFRAME
    TreeBotfrm.SetMaterialMode D3DRMMATERIAL_FROMFRAME
    'add meshes to frames
    TreeTopfrm.AddVisual TreeTopMesh
    TreeBotfrm.AddVisual TreeBotMesh
    'set color
    TreeTopfrm.SetColorRGB 0.5, 0.8, 0.1
    TreeBotfrm.SetColorRGB 0.3, 0.3, 0.2
'END_SUB TREES
'****************************************************************************************
'PRIVATE SUB BENCH_EVERYTHING
   Set BenchMesh = RMCanvas1.D3DRM.CreateMeshBuilder() 'create a mesh object
    Set Benchfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set RoadFrame = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    'load the bench
    BenchMesh.LoadFromFile "bench.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    'put the bench into the house
    Benchfrm.SetPosition Nothing, 16.5, -60, -15
    Benchfrm.SetMaterialMode D3DRMMATERIAL_FROMFRAME 'fromframe will make it be whatever color I set the frame at
    Benchfrm.SetColorRGB 0.8, 0.4, 0
    Benchfrm.AddVisual BenchMesh
'END_SUB BENCH
'**********************************************************************************************
'PRIVATE SUB CAR_EVERYTHING
    'Set carmesh = RMCanvas1.D3DRM.CreateMeshBuilder()
  'Set carfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    
    'load car
    'carmesh.LoadFromFile "car.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    'carfrm.SetPosition Nothing, 12, -57.5, -30
    'carfrm.SetMaterialMode D3DRMMATERIAL_FROMFRAME
    'carfrm.SetColorRGB 0.3, 0.3, 0.2
    'carfrm.AddVisual carmesh
    'Set CameraBoxFrame = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.CameraFrame)
'END_SUB CAR
'***********************************************************************************************
'PRIVATE SUB HOUSE_EVERYTHING
'create frames
    Set Wallfrma = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set Wallfrmb = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set Wallfrmc = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set Wallfrmd = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set Wallfrme = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    ' Now we create the boxes and add size and shape
'note that walle is actually the roof
    Set Walla = RMCanvas1.CreateBoxMesh(19.5, 5, 0.5)
    Set Wallb = RMCanvas1.CreateBoxMesh(19.5, 5, 0.5)
    Set Wallc = RMCanvas1.CreateBoxMesh(0.5, 5, -9.5)
    Set Walld = RMCanvas1.CreateBoxMesh(0.5, 5, -9.5)
    Set Walle = RMCanvas1.CreateBoxMesh(20, 1, 14)
'set wall texture,the co-ords are the same as your box size and shape
     Set texture = RMCanvas1.D3DRM.LoadTexture("WallTexture.bmp")
    Walla.SetTexture texture
    Walla.SetTextureCoordinates 19.5, 5, 0.5
    Wallb.SetTexture texture
    Wallb.SetTextureCoordinates 19.5, 5, 0.5
    Wallc.SetTexture texture
    Wallc.SetTextureCoordinates 0.5, 5, -9.5
    Walld.SetTexture texture
    Walld.SetTextureCoordinates 0.5, 5, -9.5
    'add house meshes to frames
    Wallfrma.AddVisual Walla
    Wallfrmb.AddVisual Wallb
    Wallfrmc.AddVisual Wallc
    Wallfrmd.AddVisual Walld
    Wallfrme.AddVisual Walle
   'now we make a house shape out of the walls
    Wallfrma.SetPosition Nothing, 15, -57.5, -12.5
    Wallfrmb.SetPosition Nothing, 15, -57.5, -21.5
    Wallfrmc.SetPosition Nothing, 5, -57.5, -17.5
    Wallfrmd.SetPosition Nothing, 25, -57.5, -17.5
    Wallfrme.SetPosition Nothing, 15, -55, -16.6
 'END_SUB HOUSE
'**************************************************************************************
'PRIVATE SUB BUILDING_EVERYTHING
    Set bWallfrm1 = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set bWallfrm2 = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set bWallfrm3 = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set bWallfrm4 = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    Set bRooffrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
    'create the boxes that make up the building
    Set bWall1 = RMCanvas1.CreateBoxMesh(25, 75, 0.5)
    Set bWall2 = RMCanvas1.CreateBoxMesh(25, 75, 0.5)
    Set bWall3 = RMCanvas1.CreateBoxMesh(0.05, 75, -25.5)
    Set bWall4 = RMCanvas1.CreateBoxMesh(0.05, 75, -25.5)
    Set bRoof = RMCanvas1.CreateBoxMesh(20, 1, 25)
    'set texture for the building
    Set Btexture = RMCanvas1.D3DRM.LoadTexture("btt.bmp")
    Set Btexture2 = RMCanvas1.D3DRM.LoadTexture("btt1.bmp")
    bWall1.SetTexture Btexture
    bWall1.SetTextureCoordinates 19.5, 5, 0.5
    bWall2.SetTexture Btexture
    bWall2.SetTextureCoordinates 19.5, 5, 0.5
    bWall3.SetTexture Btexture2
    bWall3.SetTextureCoordinates 0.5, 5, -9.5
    bWall4.SetTexture Btexture2
    bWall4.SetTextureCoordinates 0.5, 5, -9.5
'add meshes to frames
    bWallfrm1.AddVisual bWall1
    bWallfrm2.AddVisual bWall2
    bWallfrm3.AddVisual bWall3
    bWallfrm4.AddVisual bWall4
    bRooffrm.AddVisual bRoof
 'make a building shape
    bWallfrm1.SetPosition Nothing, -45, -55, -10
    bWallfrm2.SetPosition Nothing, -45, -55, -35.5
    bWallfrm3.SetPosition Nothing, -32.5, -55, -23
    bWallfrm4.SetPosition Nothing, -57.5, -55, -23
    bRooffrm.SetPosition Nothing, -35, -22, -29
 'END_SUB BUILDING
'**********************************************************************************
'PRIVATE SUB SURFACE_EVERYTHING
Set Surfacefrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
Set Surface = RMCanvas1.CreateBoxMesh(5000, 0.01, 5000)
Surface.GetFace(4).SetColorRGB 0, 1, 0
Surface.GetFace(5).SetColorRGB 0, 1, 0
'set texture grass texture
    'Set Stexture = RMCanvas1.D3DRM.LoadTexture("grass3.bmp")
    'Surface.SetTexture Stexture
    'Surface.SetTextureCoordinates 10, 2, 190
Surfacefrm.AddVisual Surface
Surfacefrm.SetPosition Nothing, 0, -61, 25
'END_SUB SURFACE
'**************************************************************************************
'PRIVATE SUB LOT_EVERYTHING
Set Lotfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
Set Lot = RMCanvas1.CreateBoxMesh(45, 0.2, 35)
Lotfrm.AddVisual Lot
Lotfrm.SetPosition Nothing, -42, -61, -55
Set Ltexture = RMCanvas1.D3DRM.LoadTexture("lot.bmp")
Lot.SetTexture Ltexture
Lot.SetTextureCoordinates 5, 5, 5
'END SUB
'***************************************************************************************
'PRIVATE SUB LIGHTS_EVERYTHING
'Option1 = Day and Option2 = Night 'keep this in mind
If Option1.Value = True Then Set Lightfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
If Option2.Value = True Then Set Lightfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
'If Option2.Value = True Then Set Lspotfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
If Option2.Value = True Then Set light = RMCanvas1.D3DRM.CreateLightRGB(D3DRMLIGHT_POINT, -3.2, -3.2, -2.2)
If Option1.Value = True Then Set light = RMCanvas1.D3DRM.CreateLightRGB(D3DRMLIGHT_POINT, 0.5, 0.5, 0.5)
'If Option2.Value = True Then Set Lspot = RMCanvas1.D3DRM.CreateLightRGB(D3DRMLIGHT_SPOT, 50, 50, 50)
If Option1.Value = True Then light.SetRange 50000
If Option2.Value = True Then light.SetRange 50000
'If Option2.Value = True Then Lspot.SetRange 5000000000#
'If Option2.Value = True Then Lspot.SetUmbra (1)
'If Option2.Value = True Then Lspot.GetUmbra
If Option1.Value = True Then Set Lights = RMCanvas1.CreateBoxMesh(2, 0.5, 1)
If Option2.Value = True Then Set Lights = RMCanvas1.CreateBoxMesh(2, 0.5, 1)
If Option2.Value = True Then Set Lspotm = RMCanvas1.CreateBoxMesh(-150, -150, -150)
If Option1.Value = True Then Lights.GetFace(0).SetColorRGB 0.5, 0.5, 0.5
If Option1.Value = True Then Lights.GetFace(1).SetColorRGB 0.5, 0.5, 0.5
If Option2.Value = True Then Lights.GetFace(0).SetColorRGB -3.2, -3.2, -2.2
If Option2.Value = True Then Lights.GetFace(1).SetColorRGB -3.2, -3.2, -2.2
'If Option2.Value = True Then Lspotm.GetFace(0).SetColorRGB 100, 50, 100
'If Option2.Value = True Then Lspotm.GetFace(1).SetColorRGB 100, 50, 100
If Option1.Value = True Then Lightfrm.AddLight light
If Option1.Value = True Then Lightfrm.SetPosition Nothing, 999, 1700, -999
If Option2.Value = True Then Lightfrm.AddLight light
If Option2.Value = True Then Lightfrm.SetPosition Nothing, 999, 1700, -999
'If Option2.Value = True Then Lspotfrm.AddLight Lspot
'If Option2.Value = True Then Lspotfrm.SetPosition Nothing, -42, -61, -55
If Option2.Value = True Then Set Ambfrm = RMCanvas1.D3DRM.CreateFrame(RMCanvas1.SceneFrame)
If Option2.Value = True Then Set Amblight = RMCanvas1.D3DRM.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.7, 0.7, 0.7)
If Option2.Value = True Then Amblight.SetRange 50000
If Option2.Value = True Then Set Ambm = RMCanvas1.CreateBoxMesh(-150, -150, -150)
If Option2.Value = True Then Ambm.GetFace(0).SetColorRGB 0.7, 0.7, 0.7
If Option2.Value = True Then Ambm.GetFace(1).SetColorRGB 0.7, 0.7, 0.7
If Option2.Value = True Then Ambfrm.AddLight Amblight
If Option2.Value = True Then Ambfrm.SetPosition Nothing, 45, -55, 100
'END_SUB Lights
'***********************************************************************************
'PRIVATE SUB ROAD_EVERYTHING
'make road
Set Road = RMCanvas1.CreateBoxMesh(15, 0.05, 190)
'colors for road
Road.GetFace(4).SetColorRGB 0.5, 0.5, 0.5
Road.GetFace(5).SetColorRGB 0.5, 0.5, 0.5
'add the meshes to the frames
RoadFrame.AddVisual Road
RoadFrame.SetPosition Nothing, -17.5, -60.8, 0
'END_SUB ROAD
'*********************************************************************************

    
    
   
RMCanvas1.CameraFrame.SetPosition RMCanvas1.CameraFrame, -17.5, -59.5, -80
V.z = 1
    
Do
        
DoEvents


RMCanvas1.CameraFrame.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, go
       
    RMCanvas1.Update 'update
         
    Loop


End Sub

Private Sub RMCanvas1_KeyDown(keyCode As Integer, Shift As Integer)
Dim You As D3DVECTOR 'you
Dim TreeV As D3DVECTOR 'the tree stump
Select Case keyCode
        Case vbKeyLeft
        RMCanvas1.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 4, 0, -4 * PI / 180
         'Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, -3, 0, 0)
       Case vbKeyRight
        RMCanvas1.CameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 4, 0, 4 * PI / 180
        'Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, 3, 0, 0)
        
        Case vbKeyUp
        go = Round(go, 2) + 1
        'collision detection still under development
        
        'RMCanvas1.CameraFrame.GetPosition Nothing, You 'get the camera's position
        'TreeBotfrm.GetPosition Nothing, TreeV
        'If TreeV.x - 3.3 < You.x And TreeV.x + 3.3 > You.x Then
        'If TreeV.z - 3.3 < You.z And TreeV.z + 3.3 > You.z Then
        'RMCanvas1.CameraFrame.SetPosition RMCanvas1.CameraFrame, 0, 0, -3
        'go = 0
        'End If
        'End If
        'End If
  
  'Next
        Case vbKeyDown
        Call RMCanvas1.CameraFrame.SetPosition(RMCanvas1.CameraFrame, 0, 0, -3)
        
        Case vbKeyPageUp
        'RMCanvas1.CameraFrame.AddRotation D3DRMCOMBINE_AFTER, 15, 5, -16.6, 0.02
        'UpDo = UpDo - UpDoSpeed
        Case vbKeyPageDown
        'UpDo = UpDo + UpDoSpeed
        'End If
        'End If

    End Select
'End If
End Sub

Private Sub RMCanvas1_KeyUp(keyCode As Integer, Shift As Integer)
Select Case keyCode
Case vbKeyUp
go = 0
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


