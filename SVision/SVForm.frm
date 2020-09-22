VERSION 5.00
Begin VB.Form SVForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Sterieo Vision"
   ClientHeight    =   6855
   ClientLeft      =   390
   ClientTop       =   2250
   ClientWidth     =   10110
   FillColor       =   &H00FFFFFF&
   Icon            =   "SVForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   Begin VB.Frame Frame5 
      Caption         =   "Information"
      Height          =   2415
      Left            =   7800
      TabIndex        =   33
      Top             =   4320
      Width           =   2175
      Begin VB.Label Label_Visible_Triangles 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label_Total_Triangles 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "/"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   1920
         Width           =   135
      End
      Begin VB.Label Label12 
         Caption         =   "/"
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label_Visible_Faces 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label_Total_Faces 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Total in object"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label FPS_Display 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Visible on Screen"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Faces/Triangles"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "FPS -"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Object"
      Height          =   2415
      Left            =   4800
      TabIndex        =   2
      Top             =   4320
      Width           =   2895
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   2535
         Begin VB.CheckBox Check_Lighting 
            Caption         =   "Dynamic Lighting"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton Option_Gouraud_Shading 
            Caption         =   "Gouraud Shading"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option_Flat_Shading 
            Caption         =   "Flat Shading"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.CheckBox Check_Texture 
         Caption         =   "Show Texture"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Combo_Object 
         Height          =   315
         ItemData        =   "SVForm.frx":030A
         Left            =   840
         List            =   "SVForm.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Object:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "View Adjustments"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   4575
      Begin VB.CheckBox Check_Cross 
         Caption         =   "Cross Eyed Viewing"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text_scale 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1440
         Width           =   495
      End
      Begin VB.HScrollBar HScroll_Scale 
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton Command_View_Default 
         Caption         =   "Default"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll_Depth 
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.HScrollBar HScroll_Img_Sep 
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.HScrollBar HScroll_FOV 
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text_Img_Sep 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text_Depth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text_FOV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Scale"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Image Separation"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Depth"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "FOV"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   4155
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   0
      Width           =   10095
   End
   Begin VB.Menu Menu_Application 
      Caption         =   "&Application"
      Begin VB.Menu Menu_Application_Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "&Help"
      Begin VB.Menu Menu_Help_Help 
         Caption         =   "&Help"
      End
      Begin VB.Menu Menu_Help_seporator 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Help_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "SVForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
'
' Stereo Vision
' Written By: Dennis C. Urie   quantam4@home.com
' Last Updated 1/9/2001
'
' Program Description:
'
'   This program renders two images of a simple model from separate camera
'positions.  When viewed similar to a Stereogram image it produces a TRUE 3
'dimensional view of an object.  No special viewing devices are needed to
'view in full 3D. (See the help file for how to view it properly.)
'
'   This example uses custom D3DTLVertex (Transformed and Lit Vertices) for
'rendering.  The program is responsible for the transforming and lighting
'of the 3D scene and only sends the screen coordinates, rhw, and color of
'each vertex to D3D for rendering.
'
'   This program implements most functions needed for 3D rendering in SOFTWARE
'including:
'
'  3D Object Rotation using 3X3 Matrix Transformation
'  Surface Normals
'  Backface Culling
'  3D Projection to 2D screen
'  Z Sorting
'  Dynamic Lighting (Flat Shading and Gouraud Shading)
'
'  The remaining functions (ie. color interpolation and texture mapping) are
'handled in hardware by Direct3D.
'
'
'  This program presents a few challenges not normally encountered with 3D rendering
'due to the need to display two images of a single model at different camera positions.
'
'Here are a few notes specific to this sample:
'
'      For performance and simplicity, a model is loaded directly into world
'   coordinates.  Because the model never moves there is no need for a translation
'   matrix.  Object translation into Camera space is accomplished using "Brute Force".
'
'      The camera's Y and Z coordinates are fixed.  The camera angle is fixed along
'  the world coordinates Z axis.  Though there is no code to rotate the camera angle,
'  the code does allow for any camera position along any axis.
'
'      For performance, surface and vertex normals, are calculated only when the model
'  is loaded.  Normals are then rotated in the same manner as world coordinates to ensure
'  proper backface culling and lighting calculations.
'
'      Z Ordering, Rotation, and Lighting are calculated only on the original model
'  in world coordinates.
'
'      Backface culling comparisons need to be performed with both camera positions
'  resulting in extra calculations for the 2nd camera.
'
'      Projection calculations are done with both cameras resulting in twice as many
'  projection calculations.
'
'--------------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------
' API calls for mouse position functions
'------------------------------------------
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Dim CursorPos As Point       'Holds original cursor position during object rotation
Dim newCursorPos As Point    'Used to make rotation increment calculations
Dim MouseMove As Boolean     'Toggle to enable mouse to rotate an object

'-----------------------------------------------
'Declare variables needed for Direct3D
'-----------------------------------------------
 ' Direct 3D Objects
Dim dm_DX As New DirectX8
Dim dm_D3D As Direct3D8
Dim dm_D3DDevice As Direct3DDevice8
Dim d3dx As D3DX8
Dim m_d3dpp As D3DPRESENT_PARAMETERS
Dim m_D3DCaps As D3DCAPS8
Dim m_d3ddm As D3DDISPLAYMODE
Dim d3dtTexture As Direct3DTexture8

'Flexible vertex format the describes transformed and lit vertices.
Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR


'-------------------------------------
'Custom type Definitions
'-------------------------------------
Private Const MAX_NUM_FACES = 330    'Maximum number of faces a model can have

Private Type typeMatrix_3x3
 Data(2, 2) As Single
End Type

Private Type typeWorldVertex         'Holds world oriented vertices
 X As Single                            'World X coordinante
 Y As Single                            'World Y coordinante
 Z As Single                            'World Z coordinante
End Type

Private Type typeWorldFace                 'Face Information used for world vertex
 WorldVerts(0 To 4) As typeWorldVertex        'World coordinantes
 VertexNormal(0 To 3) As typeWorldVertex      'Vertex Normals (Used for gouraud shading only)
 SurfaceNormal As typeWorldVertex             'Surface normal (Used for flat shading and backface culling)
 Color As Long                                'Color of face
 ZDist As Single                              'Average Z distance used for Z sorting
 PlaneConstant As Single                      'Plane constant used for backface culling
End Type

Private Type typeObject                    'Container used to hold all face information for a model
 Face(0 To MAX_NUM_FACES) As typeWorldFace    'Face information
 NumFaces As Integer                          'Total number of faces in an object
End Type

Private Type typeScreenVertex 'Custom 3D vertex type used to send vertex data to direct3D for rendering
 X As Single                     ' X Screen coordinante
 Y As Single                     ' Y screen coordinante
 Z As Single                     ' (Not used with Textured Lit vertexes, but needed in custom type)
 rhw As Single                   ' Used to draw textures with the correct perspective
 Color As Long                   ' Color of vertex
 specular As Long                ' Specular color
 tu As Single                    ' Horizontal texture coordinante
 tv As Single                    ' Vertical texture coordinante
End Type

Private Type typeTriangleFan                  'Vertex info array for a triangle fan
 screenVerts(0 To 3) As typeScreenVertex
End Type

Private Type typeScreenFace                   'Face array
 Face(0 To MAX_NUM_FACES) As typeTriangleFan
End Type

Dim Eye(0 To 1) As typeScreenFace       'Projected Vertex info to be sent to direct3D for rendering
Dim Camera(0 To 1) As typeWorldVertex   'Camera position for each "Eye"
Dim Object As typeObject                'The model container


'-----------------------------------------
'Projection variables
'-----------------------------------------
Private Const SCREEN_SCALE_WIDTH = 1152  'Used as starting value to compensate for screen resolutions
Dim pi As Single                         'not your mother's apple pi
Dim FOV As Single                        'field of Vision
Dim DistFromScreen As Single             'Viewer distance from screen in pixels used to caculate perspective
Dim PerspectiveCompensate As Integer     'Multiplier value to compensate for perspective
Dim DefaultScaleVal As Single            'Caculated scale value based on screen resolution
Dim ScaleVal As Single                   'User adjustable scale value
Dim EyeOffset(0 To 1) As Integer         'Holds value used to separate images during projection
Dim ImageSeparation As Long              'Used to separate the two projected images

Dim ScreenCenterX As Integer             'Center coordinantes of 3D window in pixels
Dim ScreenCenterY As Integer

Dim ShowTexture As Boolean               'Toggle to display textures


'---------------------------------------
'Z ordering variables
'---------------------------------------
Dim ZorderArray(MAX_NUM_FACES) As Integer  'Holds pointers to faces
Dim NumvisibleFaces As Integer             'Holds number of visible faces for looping during calculations

'---------------------------------------
'Dynamic Lighting variables
'---------------------------------------
Dim DynamicLight As Boolean                  'Toggle for enabling dynamic lighting
Dim LightSource As typeWorldVertex           'Holds coordinantes of light source
Private Const LIGHT_NORMALIZE_VALUE = 0.4    'Used to compensate lighting effect
Private Const RECP_LIGHT_NORMALIZE = 1 / (1 + LIGHT_NORMALIZE_VALUE) 'Recriprical of LIGHT_NORMALIZE_VALUE + 1
                                                                     
Private Const RED_INC = 65536            'Constants used for caculating colors on a per byte basis
Private Const Recp_RED_INC = 1 / 65536
Private Const GREEN_INC = 256
Private Const Recp_GREEN_INC = 1 / 256
                                                                     
                                                                     ' DO NOT CHANGE THIS VALUE!
'---------------------------------------
' Rotation variables
'---------------------------------------
Private Const ROTATE_INC = 0.01          'Rotation increments in radians

Private Const AXIS_X = &H1               'Constants to define axis for rotation
Private Const AXIS_Y = &H2
Private Const AXIS_Z = &H3

Private Sub Check_Cross_Click()
'-------------------------------------------------
' Call HScroll_Img_Sep_Change() to force update
'-------------------------------------------------

 Call HScroll_Img_Sep_Change

End Sub

Private Sub Form_Load()
'------------------------------------------------------
' Load Form, Set up control limits, Call initialization
' routines then call main program loop
'------------------------------------------------------
 Dim n As Integer

 Me.Show

 pi = 4 * Atn(1) 'caculate the value of pi

 'form and control settings that differ from default
 With Me        'This form
  .AutoRedraw = True
  .ScaleMode = 3     'Pixel               Used to relay form size in pixels to DirectX
  .BorderStyle = 1   'Fixed Single        Keeps form from being resized by user
  .Caption = "Stereo Vision"
  .KeyPreview = True '                    Allows keydown and keyup events on form
 End With
 
 App.Title = "Stereo Vision"
 App.HelpFile = App.Path & "\SVHelp.chm"
 
 With Picture1
  .ScaleMode = 3 'Pixel
 End With
  
  ' Used to position images
 ScreenCenterX = Picture1.ScaleWidth / 2
 ScreenCenterY = Picture1.ScaleHeight / 2
 
   'determine proper scale size based on screen resolution
 DefaultScaleVal = (Screen.Width / Screen.TwipsPerPixelX) / SCREEN_SCALE_WIDTH
  
 
 With HScroll_Img_Sep  'see Set_View_Defaults() for discriptions
  .Min = 32            ' of controls
  .Max = 512
  .SmallChange = 1
  .LargeChange = 5
 End With
 
 With HScroll_Depth
  .Min = 0
  .Max = 200
  .SmallChange = 1
  .LargeChange = 5
 End With
 
 Camera(0).Z = 100
 Camera(1).Z = 100
 
 With HScroll_FOV
  .Min = 5
  .Max = 75
  .SmallChange = 1
  .LargeChange = 5
 End With
  
 With HScroll_Scale
  .Min = 1
  .Max = 200
  .SmallChange = 1
  .LargeChange = 10
 End With
  
 Set_View_Defaults
  
  ' Add model names to Combo Box (These match model file names)
 Combo_Object.AddItem "Cube", 0
 Combo_Object.AddItem "Sphere", 1
 Combo_Object.AddItem "BoxFrame", 2
 Combo_Object.AddItem "Doughnut", 3
 
  'initialize Direct3D
 Call InitD3D
  'initialize Light position
 Call Init_Light
  'Call the main program loop
 Call Main_Loop
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
'----------------------------------------------
' Ends the program
'----------------------------------------------

 End   'This is needed because of the DoEvents statement in the main program loop

End Sub

Private Sub HScroll_Depth_Change()
'-------------------------------------------------
' Changes Image Separation Value and caculates
'  new Back Face Culling Threshold
'-------------------------------------------------

 Camera(0).X = (-HScroll_Depth.Value)
 Camera(1).X = HScroll_Depth.Value
 Text_Depth.Text = HScroll_Depth.Value

End Sub

Private Sub HScroll_Depth_Scroll()

 Call HScroll_Depth_Change

End Sub

Private Sub HScroll_FOV_Change()
'-------------------------------------------------------
' Changes Field Of Vision and makes calculations needed
'  to project image
'-------------------------------------------------------

 FOV = (HScroll_FOV / 180) * pi
 Text_FOV.Text = HScroll_FOV.Value
 DistFromScreen = (Picture1.ScaleWidth * 0.5) / Tan(FOV)
 PerspectiveCompensate = ScreenCenterX / Tan(FOV)

End Sub

Private Sub HScroll_FOV_Scroll()
 
 Call HScroll_FOV_Change

End Sub

Private Sub HScroll_Img_Sep_Change()
'-----------------------------------------------
' Sets distance value used to separate the two
'  images displayed on the screen
'-----------------------------------------------
 Dim CrossAdjust As Single
 Dim CrossMultiplier As Single

 ImageSeparation = HScroll_Img_Sep
 Text_Img_Sep.Text = HScroll_Img_Sep.Value
 
 If Check_Cross Then
  CrossAdjust = 213
  CrossMultiplier = -1
 Else
  CrossAdjust = 0
  CrossMultiplier = 1
 End If
 
 EyeOffset(0) = ScreenCenterX + ((((ImageSeparation - CrossAdjust) * 0.5) * (ScaleVal)) * CrossMultiplier)
 EyeOffset(1) = ScreenCenterX - ((((ImageSeparation - CrossAdjust) * 0.5) * (ScaleVal)) * CrossMultiplier)
 
End Sub

Private Sub HScroll_Img_Sep_Scroll()

 Call HScroll_Img_Sep_Change

End Sub

Private Sub HScroll_Scale_Change()
'------------------------------------------------
' Adjusts the size of the projected images
'------------------------------------------------

 ScaleVal = (HScroll_Scale.Value / 100) * DefaultScaleVal
 Text_scale.Text = HScroll_Scale.Value
 Call HScroll_Img_Sep_Change

End Sub

Private Sub HScroll_Scale_Scroll()

 Call HScroll_Scale_Change

End Sub

Private Sub Menu_Application_Exit_Click()

 Call Form_Unload(0)

End Sub

Private Sub Menu_Help_About_Click()
'---------------------------------------------------
' Display About form
'---------------------------------------------------
 
 Dim TopValue As Long
 Dim LeftValue As Long
     
 TopValue = Me.Top + ((Me.Height - About_Form.Height) / 2)   ' Center new form in main form
 LeftValue = Me.Left + ((Me.Width - About_Form.Width) / 2)
 About_Form.Top = TopValue
 About_Form.Left = LeftValue
 
 About_Form.Show vbModal
 
End Sub

Private Sub Menu_Help_Help_Click()
'--------------------------------------
' Show Help File
'--------------------------------------

 'Send F1 key to the keyboard buffer
 SendKeys "{F1}"

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'------------------------------------------------
' Set a flag to start recording mouse movements
'  and record current mouse position
'------------------------------------------------
 Dim retCode As Long

 If Not MouseMove Then
  MouseMove = True
  retCode = GetCursorPos(CursorPos)
 End If
 
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------
' Set flag to stop tracking mouse movements
'---------------------------------------------

 MouseMove = False
 
End Sub

Private Sub Main_Loop()
'-------------------------------------------------
' Main Program Loop
' - Called From Form_Load
'-------------------------------------------------
 Dim retCode As Long
 Dim MoveX As Integer
 Dim MoveY As Integer
 Dim TimeStamp As Single
 Dim FPScounter As Long
 
 Do
  
   DoEvents
 
  If MouseMove Then
   'Record Mouse movements, call rotate routine, then reset mouse to original coordinantes
   retCode = GetCursorPos(newCursorPos) 'Use API Calls to track mouse movement
   MoveX = CursorPos.X - newCursorPos.X
   MoveY = CursorPos.Y - newCursorPos.Y
  
   If MoveX Or MoveY Then
    Call Rotate(-MoveY * ROTATE_INC, MoveX * ROTATE_INC, 0)
   End If
  
   retCode = SetCursorPos(CursorPos.X, CursorPos.Y)
  End If
  
  Call Backface_Culling
  
  Call ProjectView
  
  Call Draw_Scene

  'Caculate FPS and update Info Section on the form
  FPScounter = FPScounter + 1
  If TimeStamp < Timer Then
   TimeStamp = Timer + 1
   FPS_Display.Caption = FPScounter
   FPScounter = 0
   Label_Total_Faces.Caption = Object.NumFaces
   Label_Total_Triangles.Caption = Object.NumFaces * 2
   Label_Visible_Faces.Caption = NumvisibleFaces * 2
   Label_Visible_Triangles.Caption = NumvisibleFaces * 4
  End If
  
 Loop

End Sub

Private Sub Set_View_Defaults()
'--------------------------------------------
' Sets up the default View Slider Positions
' - Called From Form_Load()
'--------------------------------------------

 With HScroll_Scale    'Scales the size of the projected images
  .Value = 100
 End With
 
 With HScroll_Img_Sep  'separates the two projected images from each other
  .Value = 325
 End With
 
 With HScroll_Depth    'Distance of each eye from center along X axis
  .Value = 55
 End With
 
 With HScroll_FOV      'Sets field of vision from center (in degrees) for projected image
  .Value = 20
 End With
 
End Sub

Private Sub Check_Lighting_Click()
'------------------------------------------------
' Sets flag to enable dynamic lighting
' and enables or disables lighting type bullets
'------------------------------------------------

 If Check_Lighting.Value Then
  DynamicLight = True
  Option_Flat_Shading.Enabled = True
  Option_Gouraud_Shading.Enabled = True
 Else
  DynamicLight = False
  Option_Flat_Shading.Enabled = False
  Option_Gouraud_Shading.Enabled = False
 End If

End Sub

Private Sub Check_Texture_Click()
'-----------------------------------------------
' Set flag to enable texture
'-----------------------------------------------

 If Check_Texture.Value Then
  ShowTexture = True
 Else
  ShowTexture = False
 End If

End Sub

Private Sub Combo_Object_Click()
'----------------------------------------------
' Call routine to load selected model
'----------------------------------------------

 Call Load_Model(Combo_Object.Text & ".def")

End Sub

Private Sub Command_View_Default_Click()
'--------------------------------------------------------
' Call routine to reset view sliders to default values
'--------------------------------------------------------

 Set_View_Defaults
 
End Sub

Private Sub InitD3D()
'-----------------------------------------
'initialize D3D Device
' - Called From Form_Load()
'-----------------------------------------
Dim DevType As CONST_D3DDEVTYPE
Dim retCode As Long

 On Error Resume Next
 Set dm_D3D = dm_DX.Direct3DCreate()                     'Create direct3d object.
 
 If Err Then
  MsgBox "Problem Initalizing Direct3D. Make sure you have DirectX 8.0 or later installed."
  End
 End If
                                                            'Get info about current display
 dm_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, m_d3ddm   'and store it in object "m_d3ddm".

 Dim d3dpp As D3DPRESENT_PARAMETERS                      'Create object used to provide
 d3dpp.Windowed = 1                                      'settings for d3d device creation.
 d3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
 d3dpp.BackBufferFormat = m_d3ddm.Format
 
 On Error Resume Next
 
 'Get Video device capabilities
 DevType = D3DDEVTYPE_HAL
 Call dm_D3D.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, m_D3DCaps)
 If Err.Number Then
  Err.Clear
  DevType = D3DDEVTYPE_REF
  Call dm_D3D.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, m_D3DCaps)
 End If
 
  'Create D3D device object using hardware.
 Set dm_D3DDevice = dm_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Picture1.hWnd, _
                                    D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
 If Err Then
  retCode = MsgBox("Unable to find compatable hardware acceleration." & vbCrLf & _
                   "Press OK for software rasterization or Cancel to quit", vbOKCancel, App.Title, 0, 0)
  If retCode = vbOK Then
    'Try creating D3D object using software rasterization
   Err.Clear
   Set dm_D3DDevice = dm_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, Picture1.hWnd, _
                                    D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
   If Err Then
    MsgBox "Error creating Direct3D divice.  Make sure your hardware is working properly."
    End
   End If
  Else 'user opted to quit
   End
  End If
 End If
  
 Set d3dx = New D3DX8   'Object used to load textures
 
 Reset_d3dDevice True   'True flag denotes that the device is being initialized for the first time
 
End Sub

Private Sub Reset_d3dDevice(Optional init As Boolean)
'-------------------------------------------------------------
' Resets the device (if not called during initialization) and
' sets rendering state
' - Called From init_D3D()
' - Called From Draw_Scene()
'-------------------------------------------------------------
 
  If Not init Then
   On Error Resume Next
    Call dm_D3DDevice.Reset(m_d3dpp)
   If Err.Number Then
    MsgBox "Direct3D Error" & Err.Number
    End
   End If
  End If
  
  On Error Resume Next
  With dm_D3DDevice
                
   'Set the vertex shader to an FVF used with transformed and lit vertex coords.
   Call .SetVertexShader(FVF)
   'Turn off lighting
   Call .SetRenderState(D3DRS_LIGHTING, 0)
   'Set the render state that uses the alpha component as the source for blending.
   Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
   'Set the render state that uses the inverse alpha component as the destination blend.
   Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        
  End With
 
  If Err.Number Then
   MsgBox "Direct3D Error " & Err.Number
   End
  End If
 
End Sub

Private Sub Load_Texture(FileName As String)
'-------------------------------------------------
' Loads a new texture
' - Called From Load_Model()
'-------------------------------------------------

 Dim bitmapFile As String
 
 On Error Resume Next
 
 Set d3dtTexture = Nothing

 bitmapFile = App.Path & "\Textures\" & FileName
 
 Set d3dtTexture = d3dx.CreateTextureFromFileEx(dm_D3DDevice, bitmapFile, _
   D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
   D3DX_FILTER_NONE, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
 
 If Err Then
  MsgBox "Error " & Err.Number & " while loading texture."
  End
 End If

End Sub

Sub Draw_Scene()
'--------------------------------------
'This draws the "Scene" to the screen
' - Called From Main_Loop()
'--------------------------------------
Dim retCode As Long
Dim n As Integer
Dim e As Integer

  'Test to see if the device is still available, if not then exit this sub and test
  ' again on the next program loop to see if it's ready to be reset
  On Local Error Resume Next
  retCode = dm_D3DDevice.TestCooperativeLevel
  If retCode = D3DERR_DEVICELOST Then           ' if it is lost then exit sub
   Exit Sub
  ElseIf retCode = D3DERR_DEVICENOTRESET Then   ' if it's ready to be reset then do so
   Reset_d3dDevice
  End If
  
 dm_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0&, 1#, 0    ' Clear the backbuffer
  
 dm_D3DDevice.BeginScene                                        ' Begin rendering process
    
 If ShowTexture Then
  dm_D3DDevice.SetTexture 0, d3dtTexture      'Set texture to use with rendering
 Else
  dm_D3DDevice.SetTexture 0, Nothing          'or set to nothing to eliminate memory leaks
 End If
    
   'Draw all of visible triangle strips
 With Object
   'Loop through each visible face
  For n = 1 To NumvisibleFaces
    'loop for each "eye"
   For e = 0 To 1
    dm_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 2, Eye(e).Face(ZorderArray(n)).screenVerts(0), Len(Eye(e).Face(ZorderArray(n)).screenVerts(0))
   Next
  Next
 End With
 
 dm_D3DDevice.EndScene                                        ' End Rendering process
 
 dm_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0            ' Display scene to the screen
    
End Sub

Private Sub Load_Model(DefFile As String)
'---------------------------------------------------------
' Loads a new model into memory and assigns the vertexes
' to the world coordinantes.
' - Called From Combo_Object_Click()
'---------------------------------------------------------

 Dim FileName As String
 Dim n As Integer
 Dim i As Integer
 Dim e As Integer
 Dim NumVerts As Integer
 Dim X As Single
 Dim Y As Single
 Dim Z As Single
 Dim nX As Single
 Dim nY As Single
 Dim nZ As Single
 Dim tu As Single
 Dim tv As Single
 Dim Color As Long
 Dim Texture As String
 Dim NumFaces As Integer

 FileName = App.Path & "\Models\" & DefFile

 On Error GoTo Problem
 
 Open FileName For Input As #1
 
 Input #1, Texture
 
 With Object
 
  Input #1, NumFaces
  .NumFaces = NumFaces
  
  'initialize zorder array
  For n = 1 To NumFaces
   ZorderArray(n) = n
  Next
  
  For n = 1 To NumFaces
   Input #1, NumVerts
   For i = 0 To NumVerts - 1
    Input #1, X, Y, Z, nX, nY, nZ, tu, tv, Color
    With .Face(n)
     'assign worldvertex data
     .WorldVerts(i).X = X
     .WorldVerts(i).Y = Y
     .WorldVerts(i).Z = Z
     .VertexNormal(i).X = nX
     .VertexNormal(i).Y = nY
     .VertexNormal(i).Z = nZ
     .Color = Color
     For e = 0 To 1
      'assign texture coordinantes to screen vertex's now since these need no calculations
      Eye(e).Face(n).screenVerts(i).tu = tu
      Eye(e).Face(n).screenVerts(i).tv = tv
     Next
    End With
   Next
  Next
 
 End With
 
Close #1
 
 'initialize surface normals
Call Init_Surface_Normals

 'load the model's texture
Call Load_Texture(Texture)

Exit Sub
Problem:

Close #1
 MsgBox "Error Loading Model Definition File"
 End

End Sub

Private Sub Init_Light()
'-----------------------------------
' initialize the light source
' - Called From Form_Load()
'-----------------------------------

 Dim Normalizer As Single

'Define a point that denotes where light is coming from
'(the light will point towards the center of the world coordinantes)
  LightSource.X = 50
  LightSource.Y = 100
  LightSource.Z = 400
    
'Normalize the light coordinantes for use with lighting calculations
  Normalizer = Sqr((LightSource.X * LightSource.X) + (LightSource.Y * LightSource.Y) + (LightSource.Z * LightSource.Z))
  LightSource.X = LightSource.X / Normalizer
  LightSource.Y = LightSource.Y / Normalizer
  LightSource.Z = LightSource.Z / Normalizer

End Sub

Private Sub Init_Surface_Normals()
'------------------------------------------------------------------
' Caculate surface normals to be used for backface culling and
'  flat light shading.  Store as Object.SurfaceNormal
' ' - Called From Load_Model()
'------------------------------------------------------------------

 Dim n As Integer
 Dim Side1 As typeWorldVertex
 Dim Side2 As typeWorldVertex
 Dim Normalizer As Single

 On Error GoTo Problem
 
 With Object
  For n = 1 To .NumFaces
   With .Face(n)
     
     'Caculate Surface normals for lighting calculations

    Side1.X = (.WorldVerts(1).X - .WorldVerts(0).X)
    Side1.Y = (.WorldVerts(1).Y - .WorldVerts(0).Y)
    Side1.Z = (.WorldVerts(1).Z - .WorldVerts(0).Z)
   
    Side2.X = (.WorldVerts(2).X - .WorldVerts(1).X)
    Side2.Y = (.WorldVerts(2).Y - .WorldVerts(1).Y)
    Side2.Z = (.WorldVerts(2).Z - .WorldVerts(1).Z)
    
    .SurfaceNormal.X = (Side1.Y * Side2.Z) - (Side1.Z * Side2.Y)
    .SurfaceNormal.Y = (Side1.Z * Side2.X) - (Side1.X * Side2.Z)
    .SurfaceNormal.Z = (Side1.X * Side2.Y) - (Side1.Y * Side2.X)
    
    Normalizer = Sqr((.SurfaceNormal.X * .SurfaceNormal.X) + _
                     (.SurfaceNormal.Y * .SurfaceNormal.Y) + _
                     (.SurfaceNormal.Z * .SurfaceNormal.Z))
    .SurfaceNormal.X = .SurfaceNormal.X / Normalizer
    .SurfaceNormal.Y = .SurfaceNormal.Y / Normalizer
    .SurfaceNormal.Z = .SurfaceNormal.Z / Normalizer
    
     'Caculate the dot product for one of the verts for use with backface culling
     
     .PlaneConstant = (.WorldVerts(0).X * .SurfaceNormal.X) + _
                      (.WorldVerts(0).Y * .SurfaceNormal.Y) + _
                      (.WorldVerts(0).Z * .SurfaceNormal.Z)
     
   End With
  Next
 End With
Exit Sub

Problem:
 
 MsgBox "Model Definition File invalad or contains errors"
 End

End Sub

Private Sub Rotate(AngleX As Single, AngleY As Single, AngleZ As Single)
'----------------------------------------------------
'Rotate all world vertices, surface normals and
'vertex normals
' - Called From Main_Loop()
'----------------------------------------------------
 Dim n As Integer
 Dim i As Integer
 Dim tempMatrix As typeMatrix_3x3
 Dim RotateMatrix As typeMatrix_3x3

 Call Matrix_Set_Identity(tempMatrix)      'Create temporary Identity Matrix
 Call Matrix_Set_Identity(RotateMatrix)    'The "Master" matrix
 
 If AngleX <> 0 Then                       'Create an X rotation matrix from the
  Call Matrix_RotateX(tempMatrix, AngleX)  'identity matrix and assign new matrix
  RotateMatrix = tempMatrix                'assign new matrix to a "master" matrix
 End If

 If AngleY <> 0 Then                         'Reset Temp matrix back to identity matrix,
  Call Matrix_Set_Identity(tempMatrix)       'Create an Y rotation matrix, then concentrate
  Call Matrix_RotateY(tempMatrix, AngleY)    'it with the "master" rotation matrix
  RotateMatrix = Matrix_Concentrate(RotateMatrix, tempMatrix)
 End If
 
 If AngleZ <> 0 Then                       'Repeat the process for Z rotations
  Call Matrix_Set_Identity(tempMatrix)
  Call Matrix_RotateZ(tempMatrix, AngleZ)
  RotateMatrix = Matrix_Concentrate(RotateMatrix, tempMatrix)
 End If
 
 'Rotate all Worldverts, Surface normals, and vertex normals using the "Master" matrix
 With Object
  For n = 1 To .NumFaces
   With .Face(n)
    .SurfaceNormal = Matrix_Multiply(.SurfaceNormal, RotateMatrix)
    For i = 0 To 3
     .WorldVerts(i) = Matrix_Multiply(.WorldVerts(i), RotateMatrix)
     .VertexNormal(i) = Matrix_Multiply(.VertexNormal(i), RotateMatrix)
    Next
   End With
  Next
 End With
 
End Sub

Private Function Matrix_Set_Identity(Matrix As typeMatrix_3x3)
'----------------------------------
' Creates a 3x3 identity matrix
'----------------------------------
'      0      1     2
'    +-----+-----+-----+
' 0  |  1  |  0  |  0  |
'    +-----+-----+-----+
' 1  |  0  |  1  |  0  |
'    +-----+-----+-----+
' 2  |  0  |  0  |  1  |
'    +-----+-----+-----+
'

 With Matrix

  .Data(0, 0) = 1: .Data(0, 1) = 0: .Data(0, 2) = 0
  .Data(1, 0) = 0: .Data(1, 1) = 1: .Data(1, 2) = 0
  .Data(2, 0) = 0: .Data(2, 1) = 0: .Data(2, 2) = 1

 End With

End Function

Private Function Matrix_RotateX(Matrix As typeMatrix_3x3, Angle As Single)
'------------------------------------------
' Assigns values needed to rotate around
' the X axis.
'------------------------------------------
'      0      1     2
'    +-----+-----+-----+
' 0  |     |     |     |
'    +-----+-----+-----+
' 1  |     | COS | Sin |
'    +-----+-----+-----+
' 2  |     |-SIN | Cos |
'    +-----+-----+-----+

 With Matrix
  .Data(1, 1) = Cos(Angle)
  .Data(2, 2) = Cos(Angle)
  .Data(2, 1) = -Sin(Angle)
  .Data(1, 2) = Sin(Angle)
 End With

End Function

Private Function Matrix_RotateY(Matrix As typeMatrix_3x3, Angle As Single)
'------------------------------------------
' Assigns values needed to rotate around
' the X axis
'------------------------------------------
'      0      1     2
'    +-----+-----+-----+
' 0  | COS |     |-SIN |
'    +-----+-----+-----+
' 1  |     |     |     |
'    +-----+-----+-----+
' 2  | SIN |     | COS |
'    +-----+-----+-----+
 
 With Matrix
  .Data(0, 0) = Cos(Angle)
  .Data(0, 2) = -Sin(Angle)
  .Data(2, 0) = Sin(Angle)
  .Data(2, 2) = Cos(Angle)
 End With
 
End Function

Private Function Matrix_RotateZ(Matrix As typeMatrix_3x3, Angle As Single)
'------------------------------------------
' Assigns values needed to rotate around
' the X axis
'------------------------------------------
'      0      1     2
'    +-----+-----+-----+
' 0  | COS | SIN |     |
'    +-----+-----+-----+
' 1  |-SIN | COS |     |
'    +-----+-----+-----+
' 2  |     |     |     |
'    +-----+-----+-----+
 
 With Matrix
  .Data(0, 0) = Cos(Angle)
  .Data(1, 1) = Cos(Angle)
  .Data(0, 1) = Sin(Angle)
  .Data(1, 0) = -Sin(Angle)
 End With

End Function

Private Function Matrix_Concentrate(MatrixA As typeMatrix_3x3, MatrixB As typeMatrix_3x3) As typeMatrix_3x3
'--------------------------------------------
' Concentrates two matrices into one matrix
'--------------------------------------------
 With Matrix_Concentrate
 
  .Data(0, 0) = (MatrixA.Data(0, 0) * MatrixB.Data(0, 0)) + _
                (MatrixA.Data(0, 1) * MatrixB.Data(1, 0)) + _
                (MatrixA.Data(0, 2) * MatrixB.Data(2, 0))

  .Data(0, 1) = (MatrixA.Data(0, 0) * MatrixB.Data(0, 1)) + _
                (MatrixA.Data(0, 1) * MatrixB.Data(1, 1)) + _
                (MatrixA.Data(0, 2) * MatrixB.Data(2, 1))

  .Data(0, 2) = (MatrixA.Data(0, 0) * MatrixB.Data(0, 2)) + _
                (MatrixA.Data(0, 1) * MatrixB.Data(1, 2)) + _
                (MatrixA.Data(0, 2) * MatrixB.Data(2, 2))

  .Data(1, 0) = (MatrixA.Data(1, 0) * MatrixB.Data(0, 0)) + _
                (MatrixA.Data(1, 1) * MatrixB.Data(1, 0)) + _
                (MatrixA.Data(1, 2) * MatrixB.Data(2, 0))

  .Data(1, 1) = (MatrixA.Data(1, 0) * MatrixB.Data(0, 1)) + _
                (MatrixA.Data(1, 1) * MatrixB.Data(1, 1)) + _
                (MatrixA.Data(1, 2) * MatrixB.Data(2, 1))

  .Data(1, 2) = (MatrixA.Data(1, 0) * MatrixB.Data(0, 2)) + _
                (MatrixA.Data(1, 1) * MatrixB.Data(1, 2)) + _
                (MatrixA.Data(1, 2) * MatrixB.Data(2, 2))

  .Data(2, 0) = (MatrixA.Data(2, 0) * MatrixB.Data(0, 0)) + _
                (MatrixA.Data(2, 1) * MatrixB.Data(1, 0)) + _
                (MatrixA.Data(2, 2) * MatrixB.Data(2, 0))

  .Data(2, 1) = (MatrixA.Data(2, 0) * MatrixB.Data(0, 1)) + _
                (MatrixA.Data(2, 1) * MatrixB.Data(1, 1)) + _
                (MatrixA.Data(2, 2) * MatrixB.Data(2, 1))

  .Data(2, 2) = (MatrixA.Data(2, 0) * MatrixB.Data(0, 2)) + _
                (MatrixA.Data(2, 1) * MatrixB.Data(1, 2)) + _
                (MatrixA.Data(2, 2) * MatrixB.Data(2, 2))
                
 End With

End Function

Private Function Matrix_Multiply(Vertex As typeWorldVertex, Matrix As typeMatrix_3x3) As typeWorldVertex
'----------------------------------------------------
' Caculates new "rotated" coordinantes for a vertex
' based on the supplied matrix
'----------------------------------------------------

 With Matrix_Multiply
 
  .X = (Matrix.Data(0, 0)) * Vertex.X + _
       (Matrix.Data(1, 0)) * Vertex.Y + _
       (Matrix.Data(2, 0)) * Vertex.Z
  
  .Y = (Matrix.Data(0, 1)) * Vertex.X + _
       (Matrix.Data(1, 1)) * Vertex.Y + _
       (Matrix.Data(2, 1)) * Vertex.Z
  
  .Z = (Matrix.Data(0, 2)) * Vertex.X + _
       (Matrix.Data(1, 2)) * Vertex.Y + _
       (Matrix.Data(2, 2)) * Vertex.Z
 
 End With

End Function

Private Sub Backface_Culling()
'----------------------------------------------------------
' determine if a face is facing towards the viewer
' and store a pointer to that face in ZorderArray().
' - Called From Main_Loop
'----------------------------------------------------------

Dim n As Integer
Dim i As Integer
Dim FaceCount As Integer
Dim EyeSourceDot As Single
Dim ViewerPlane(0 To 1) As Single
  
   'Note about backface culling:  When using Transformed and Lit Vertices and triangle fans,
   ' Direct3D only draws Triangles where the vertices are oriented in a clockwise fashion.
   ' If you were to rotate a triangle so it's face is pointing away from the viewer the vertices
   ' are now oriented counter-clockwise, and therefore not drawn hence, D3d has built in
   ' backface culling.
   '
   '  Even though Direct3d performs it's own backface culling on the projected image, a
   ' substantial increase in performance can be gained by NOT performing Z sorting,
   ' lighting, and projection calculations on faces that are not visible.
   
   '  To determine if a face is pointing towards or away from the viewer, compare the face's
   ' Plane constant to the point where the viewer is located.
 
   ' Object.Face(n).PlaneConstant is caculated when the model is loaded in Init_Surface_Normals()
 
 With Object
  For n = 1 To .NumFaces
   With .Face(n)
    
    'Caculate the plane constant for each camera's viewer position on the same orientation
    ' as the face's surface normal
    For i = 0 To 1
     ViewerPlane(i) = (Camera(i).X * .SurfaceNormal.X) + _
                      (Camera(i).Y * .SurfaceNormal.Y) + _
                      ((Camera(i).Z + DistFromScreen) * .SurfaceNormal.Z)
    Next
    
    'Check to see if either viewer position is in front of the face
    If ViewerPlane(0) >= .PlaneConstant Or _
       ViewerPlane(1) >= .PlaneConstant Then
     
     'If so then add a pointer to that face in ZorderArray
     FaceCount = FaceCount + 1
     ZorderArray(FaceCount) = n
    End If
   End With
  Next
 End With
 
 'store total number of visible faces for use when looping through other calculations
 NumvisibleFaces = FaceCount

End Sub

Private Sub ProjectView()
'---------------------------------------------------------------
' Project view into 2D screen coordinantes
' - Called From Main_Loop()
'---------------------------------------------------------------

 Dim n As Integer
 Dim i As Integer
 Dim e As Integer
 Dim DistX As Single
 Dim DistY As Single
 Dim DistZ As Single
 Dim ScreenX As Single
 Dim ScreenY As Single
 Dim W As Integer
 Dim RicpDistZ As Single
 Dim DistZtot As Single
  
 'This routine renders two separate images from one world object.  Each image is rendered
 ' from a separate camera position providing the correct perspective for eace "eye" and
 ' stored as TLVertex in the user type eye()
  
 With Object
  For n = 1 To NumvisibleFaces    'Only loop through the faces that have been marked
   With .Face(ZorderArray(n))     ' as visible by Backface_Culling()
    
    For i = 0 To 3                'Loop through each faces vertices
    
      'Each "eye" shares the same Z distance so caculate it prior to looping through each
      ' "eye" then store it for use with Z ordering
     
     DistZ = Camera(0).Z - .WorldVerts(i).Z + DistFromScreen
     DistZtot = DistZtot + DistZ

     RicpDistZ = 1 / DistZ  'Caculate the Reciprocal of the Z distance here to help
                            'limit processing the during next loop
     
     For e = 0 To 1               'loop for each "eye"
                                 
       'Use "Brute force" to transform world vertex coordinantes to camera coordinantes
      DistX = Camera(e).X - .WorldVerts(i).X
      DistY = Camera(e).Y - .WorldVerts(i).Y
      
       'Transform image to screen coordinantes.
      ScreenX = (DistX * PerspectiveCompensate) * RicpDistZ * ScaleVal
      ScreenY = (DistY * PerspectiveCompensate) * RicpDistZ * ScaleVal
      
       'Center the image to the correct area on the screen.
      Eye(e).Face(ZorderArray(n)).screenVerts(i).X = EyeOffset(e) + ScreenX
      Eye(e).Face(ZorderArray(n)).screenVerts(i).Y = ScreenCenterY + ScreenY
      
       'Caculate rhw value
      Eye(e).Face(ZorderArray(n)).screenVerts(i).rhw = DistFromScreen * RicpDistZ
    
       'Assign Color value to TLVertex here if not using dynamic lighting.
      If Not DynamicLight Then
       Eye(e).Face(ZorderArray(n)).screenVerts(i).Color = .Color
      End If
     Next
    Next
    
    'Caculate Average Z Distance for each Face
    .ZDist = DistZtot * 0.25
    DistZtot = 0
   End With
  Next
 End With
 
  'Call lighting routines if needed
 If DynamicLight Then
  If Option_Gouraud_Shading.Value Then
   Call Gouraud_Shading
  Else
   Call Flat_Shading
  End If
 End If

  'Call Routine to sort faces by their avrage Z distance
 Call Z_Sort(1, NumvisibleFaces)

End Sub

Private Sub Flat_Shading()
'-----------------------------------------
' Caculates lighting using Flat Shading
' - Called From Project_View()
'-----------------------------------------

 Dim n As Integer
 Dim i As Integer
 Dim e As Integer
 Dim Shade As Single
 Dim R As Long
 Dim G As Long
 Dim B As Long
 Dim RGB As Long
 Dim Temp As Long
 
 With Object
  For n = 1 To NumvisibleFaces
   With .Face(ZorderArray(n))
    
    'Caculate Light shade value from surface normals
    'Shade value needs to be within the range of 0 to 1
      
    'This code caculates the SIN of the angle between the light source and the
    ' face to be shaded (a range between (-1) to 1)
    Shade = ((.SurfaceNormal.X * LightSource.X) + _
             (.SurfaceNormal.Y * LightSource.Y) + _
             (.SurfaceNormal.Z * LightSource.Z))
    
    'The shade value then needs to be normalized before caculating light values.
    
    Shade = (Shade + LIGHT_NORMALIZE_VALUE) * RECP_LIGHT_NORMALIZE
     
     ' For a more realistic lighting effect the shade value was actually normalized to a range
     ' between (-1) + LIGHT_NORMALIZE_VALUE to 1.  Since the lowest valid shade value = 0, we
     ' need to compensate for it.

    If Shade < 0 Then
     Shade = 0
    End If
            
    'Set new RGB values based on value of Shade
     
     RGB = .Color
    
      'Break the long integer color value down to red, green, and blue "Byte" values
     R = Int(RGB * Recp_RED_INC)
     RGB = RGB - (R * RED_INC)
     G = Int(RGB * Recp_GREEN_INC)
     B = RGB - (G * GREEN_INC)
    
      'Adjust individual colors accordingly
     R = Int(R * Shade)
     G = Int(G * Shade)
     B = Int(B * Shade)
    
      'Assign back to a long integer value
     RGB = (R * RED_INC) + (G * GREEN_INC) + B
   
    'assign new RGB value to all verts for current Face
    For e = 0 To 1
     For i = 0 To 3
      Eye(e).Face(ZorderArray(n)).screenVerts(i).Color = RGB
     Next
    Next
   End With
  Next
 End With

End Sub

Private Sub Gouraud_Shading()
'-----------------------------------------------
' Caculates lighting using gouraud shading
' - Called From Project_View
'-----------------------------------------------
 
 Dim n As Integer
 Dim i As Integer
 Dim e As Integer
 Dim Shade As Single
 Dim R As Long
 Dim G As Long
 Dim B As Long
 Dim RGB As Long
 
  'This is similar to Fat_Shading() except we need the loop to caculate lighting
  ' values for each vertex normal, rather than each surface normal.
 
  'The vertex normals used for the light calculations are NOT true world vertex normals,
  ' They are defined by the model definition file to assure that touching vertices have
  ' the same vertex normal and are relative to the actual direction that the face is facing.
  ' This allows for convex objects like the doughnut model to light properly.
 
 With Object
  For n = 1 To NumvisibleFaces
   With .Face(ZorderArray(n))
    
    For i = 0 To 3
     With .VertexNormal(i)
    
     'Caculate Light shade value from vertex normals
      Shade = ((.X * LightSource.X) + (.Y * LightSource.Y) + (.Z * LightSource.Z))
      Shade = (Shade + LIGHT_NORMALIZE_VALUE) * RECP_LIGHT_NORMALIZE
     
     End With
    
     If Shade < 0 Then
      Shade = 0
     End If
            
      'Set new RGB values based on value of Shade
     
     RGB = .Color
    
     R = Int(RGB * Recp_RED_INC)
     RGB = RGB - (R * RED_INC)
     G = Int(RGB * Recp_GREEN_INC)
     B = RGB - (G * GREEN_INC)
    
     R = Int(R * Shade)
     G = Int(G * Shade)
     B = Int(B * Shade)
    
     RGB = (R * RED_INC) + (G * GREEN_INC) + B
    
      'assign new RGB value to current vertex
     For e = 0 To 1
      Eye(e).Face(ZorderArray(n)).screenVerts(i).Color = RGB
     Next
    Next
   End With
  Next
 End With
 
End Sub

Private Sub Z_Sort(ByVal LowerBound As Integer, ByVal UpperBound As Integer)
'--------------------------------------------------------
' Sort faces by their average Z distance (Caculated in
' Project_View()) using a QuickSort routine.
' - Called From Project_View
'--------------------------------------------------------

'Note: While this routine is sorting based on a face's average Z distance it is only
'swapping pointers to the faces rather than the face data itself.

'Face pointers are stored in ZorderArray() and face data is contained in
'Object.Face(n) where n is replaced by the value in ZorderArray()

 Dim StopHere As Boolean
 Dim Left As Integer
 Dim Right As Integer
 Dim Temp As Integer
 Dim RefVal As Single
   
   Left = LowerBound
   Right = UpperBound
   RefVal = Object.Face(ZorderArray(Left)).ZDist
 
   While Left < Right
  
    While Object.Face(ZorderArray(Right)).ZDist < RefVal
     Right = Right - 1
    Wend
    
    While Object.Face(ZorderArray(Left)).ZDist > RefVal
     Left = Left + 1
    Wend
    
    If Left < Right Then
     Temp = ZorderArray(Left)
     ZorderArray(Left) = ZorderArray(Right)
     ZorderArray(Right) = Temp
     Right = Right - 1
    End If
        
   Wend
   
   If Right > LowerBound Then
    Call Z_Sort(LowerBound, Right)
   End If
   
   If Right + 1 < UpperBound Then
    Call Z_Sort(Right + 1, UpperBound)
   End If

End Sub

