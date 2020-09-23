VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//===========================================================
'JohnaDX7 engine version 1.03
'
' JohnaDX7 engine LandDemo2
'
'  the main engine have be updated
'  -Texture surface can be created from BMP,JPG,JPEG,GIF,TGA,PCX files
'  -Total control to the camera
'
'
'New MDL class added in order to load MDL and Quake2 character and object
'  -MDL
'    -play Mdl animation
'    -update sequence and controle sequence to animate legs,torse,mouth
'    -Scale,rotate,move any object
'    -AAbounding box colision detection will be enable to the next version
'    -getallsequence name
'  -MD2
'    -Load quake2 MD2 file
'    -Scale,rotate,move any object
'    -Load the main Texture to object
'    -Quake3 object will be supported in the next Version
'Sound engine improved
'   -Midi file support with directMusic
'   -Midi volume control and play repeatly the midi
'   -DirectMusic object accessible in order to loas segment
'
'AI engine for vehicules
'   -Now give realistic behavior to your vehicules movements
'   -Define an area to be visited by your vehicule
'   ........
'THe particle engine Have been updated
'   -Smoke
'   -Rain
'   -Snow
'   -Fire
'   -Explosion
'   ..........
'  Change the emiter to do great FX
'
'
'The Xfile Class have been Updated
'   -beter colision detection
'   -Scale,rotate,move object
'
'
'
'
'
'For any comment or suggestion
'  write me at Johna.pop@caramail.com
'
'
'
'
'
'//===============================================================










'Main engine
Dim DX7 As New johna_DX7


Dim MATH As New cMATH



'sky object
Dim SKY As New cJohna_Sphere

'temp Vector
Dim VC As D3DVECTOR
Dim Zangle
Const SKY_COLOR = &HDFDEDEF


'Landscape engine
Dim Land As New cJohna_landScape


'Halflife MDL file sample
Dim MDL As New cJohna_MDLfile
'AI for exploring randomely the terrain
Dim MDL_AI As New cJohna_AI
Dim MDL_pos As D3DVECTOR
Dim MDL_smokE As New cJohna_Particle



'Sound engine
Dim SOUND As New cJohna_Sound


Private Sub Form_Load()
Me.Refresh
Me.Show
DX7.INIT_ShowDialog Me.hWnd     'init the 3d engine

Call Me.INIT_object
Call GAME_LOOP     'call the game loop
End Sub


Sub INIT_object()


'init land
Land.Initialize DX7
Land.LoadTerrain App.Path + "\data\heightmap.jpg", 10, 0.5, 14, 14, App.Path + "\data\DIRT.jpg"

'The skyDome
SKY.InitSPhere DX7, App.Path + "\data\sky02.jpg", NORMAL_QUALITY, 8100, 8160, 8100



'Load MDLfile
MDL.Load_MDL App.Path + "\Data\chopper.mdl", DX7

MDL_AI.INIT_Ai MATH.Vector(0, 0, 0), MATH.Vector((Rnd - Rnd) * 800, 0, (Rnd - Rnd) * 800), 1.5, 0.25, 4, 0, 0, 2, 0.001, 0.01
Dim I
Randomize Timer
'prapare different locations to be visited by the copter
'location generate randomly
For I = 0 To 50
  MDL_AI.AI_Add_Target MATH.Vector((Rnd - Rnd) * 800, 0, (Rnd - Rnd) * 800)
Next I

'prepare the smoke effect
MDL_smokE.Init DX7, MATH.Vector(0, 0, 0), App.Path + "\data\smoke.bmp", 50, 1 / 5, Johna_SMOKE


'sound preparation
SOUND.INIT_SoundENGINE Me.hWnd, 200

'load a midi
SOUND.DMusic_OpenMidi App.Path + "\data\loop_midi.mid"
SOUND.DMusic_PlayMidi 1


End Sub




Sub GAME_LOOP()

'=======Place the camera actor EYE the current orientation is the Degree=0

'==============SKY object initialization
DX7.SET_camera MATH.Vector(0, 10, 0)
DX7.device_CLEAR_COLOR = SKY_COLOR



Land.vPosition = MATH.Vector(-1500, 0, -1500)
Land.Set_Scale 40, 15, 40


MDL.SetScale 1, 1, 1
MDL.SetAnimation 4
MDL.SetSpeed 1

'set the backbuffer color for drawing to the screen
DX7.BAK.SetForeColor RGB(0, 25, 255)


Do

  DoEvents
 

  'DoEvents
  If DX7.GetKEY(Johna_KEY_ESCAPE) Then GoTo END_it
  'checkeys
  Call Me.KEY_check
  
  
  'Render 3d
  DX7.D3D_DEV.BeginScene
  DX7.Clear_3D
  
  'uncomment it to add fog
  'DX7.SetFog 1, 1 / 5800, SKY_COLOR, D3DFOG_EXP2
  
  'render land
  Land.RenderAll 0
  
  
  'uncomment it to render the skydome
  'render the sky
  SKY.Render DX7
  
  
  
  
  'render MDL
  RenderMDL
  DX7.D3D_DEV.EndScene
  'end rendering

  
  DX7.BAK.DrawText 10, 40, "FPS=" + Str(DX7.FramesPerSec), 0
  DX7.BAK.DrawText 10, 60, "PRESS SPACE to locate The Copter", 0
  
  
  
  DX7.FLIPP Me.hWnd
Loop


END_it:


'free ressource
SOUND.Engine_Free
DX7.FreeDX Me.hWnd
End



End Sub




Sub RenderMDL()

'update the MDL artificial inteligence Object
MDL_AI.Update_AI
VC = MDL_AI.Get_location

'for smoke particle

'test colision with terrain
VC.y = Land.Get_Altitude_EX(VC) + 450



'update MDL location
 MDL.SetPosition VC.x, VC.y, VC.z
 MDL_pos = VC
 MDL.SetRotation 90, MDL_AI.GetAngle_DEG - 90, 0
 
 'Draw MDL_frame
 DX7.SetFog 0
 MDL.Render True
 MDL_smokE.Render DX7
 VC.y = VC.y - 380
 MDL_smokE.SetPosition VC
 
 'for soundFX
 
 'simulate soundFX moving
 'the formula: get the absolute distance between playerpos and copter's
 'pass a percentage parameter to the Directmusic volume
 SOUND.Dmusic_volume = -((((MATH.VDist(DX7.GET_CameraEYE, VC))) / 100) * 10) + 100

End Sub




Sub KEY_check()

VC = DX7.GET_CameraEYE

    VC.y = Land.Get_Altitude_EX(VC) + 80
    DX7.SET_camera VC
    

If DX7.GetKEY(Johna_KEY_UP) Then
  DX7.Camera_Move_Foward 1
 
  'SOUND.Play_BUF PLAYER_STEP
 
End If
  
If DX7.GetKEY(Johna_KEY_RCONTROL) Then
     DX7.Camera_Move_Foward 8
     'SOUND.Play_BUF PLAYER_STEP
    
End If

If DX7.GetKEY(Johna_KEY_DOWN) Then
  
     'SOUND.Play_BUF PLAYER_STEP
     DX7.Camera_Move_Backward 1
End If

If DX7.GetKEY(Johna_KEY_LEFT) Then _
  DX7.Camera_Move_Left 0.0005

If DX7.GetKEY(Johna_KEY_RIGHT) Then _
  DX7.Camera_Move_Right 0.0005


If DX7.GetKEY(Johna_KEY_NUMPAD8) Then _
  DX7.Camera_Move_UP 0.0005

If DX7.GetKEY(Johna_KEY_NUMPAD2) Then _
  DX7.Camera_Move_DOWN 0.0005



If DX7.GetKEY(Johna_KEY_NUMPAD4) Then _
  DX7.CAM_step_LEFT 1

If DX7.GetKEY(Johna_KEY_NUMPAD6) Then _
  DX7.CAM_step_RIGHT 1


If DX7.GetKEY(Johna_KEY_SPACE) Then
  VC = MDL_pos
  VC.y = VC.y - 350
  DX7.SET_camera VC

End If

If DX7.GetKEY(Johna_KEY_ADD) Then DX7.Camera_Elevator_UP 1
If DX7.GetKEY(Johna_KEY_SUBTRACT) Then DX7.Camera_Elevator_DOWN 1





End Sub






