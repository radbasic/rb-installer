VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmIniSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RAD Basic Installer"
   ClientHeight    =   4752
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8436
   Icon            =   "FrmIniSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4752
   ScaleWidth      =   8436
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmDstFolder 
      BorderStyle     =   0  'None
      Height          =   3372
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   8172
      Begin VB.CommandButton CmdChangePath 
         Caption         =   "Change..."
         Height          =   372
         Left            =   5520
         TabIndex        =   16
         Top             =   360
         Width           =   1332
      End
      Begin VB.TextBox TxtDstFolder 
         Height          =   288
         Left            =   120
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   480
         Width           =   5292
      End
      Begin VB.Label LblDstFolder 
         Caption         =   "Install RAD Basic to:"
         Height          =   252
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2532
      End
   End
   Begin VB.Frame FrameIni 
      BorderStyle     =   0  'None
      Height          =   3012
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   8172
      Begin VB.Frame FrmVersion 
         Height          =   1692
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   7812
         Begin VB.CommandButton CmdRemove 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   372
            Left            =   6120
            TabIndex        =   32
            Top             =   1200
            Width           =   1452
         End
         Begin VB.CommandButton CmdModify 
            Caption         =   "Modify"
            Enabled         =   0   'False
            Height          =   372
            Left            =   6120
            TabIndex        =   31
            Top             =   720
            Width           =   1452
         End
         Begin VB.CommandButton CmdUpdate 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   372
            Left            =   6120
            TabIndex        =   30
            Top             =   240
            Width           =   1452
         End
         Begin VB.Label Label8 
            Caption         =   "Already updated"
            Height          =   252
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   1572
         End
         Begin VB.Label Label7 
            Caption         =   "RAD Basic nightly"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   10.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   2652
         End
      End
      Begin VB.CommandButton CmdInstallNew 
         Caption         =   "New install"
         Height          =   612
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   3492
      End
      Begin VB.Label LblNoVersions 
         Caption         =   "No versions detected."
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   5052
      End
   End
   Begin VB.Frame FrmInstallng 
      BorderStyle     =   0  'None
      Height          =   3252
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   8172
      Begin MSComctlLib.ProgressBar PgbSetup 
         Height          =   252
         Left            =   720
         TabIndex        =   26
         Top             =   1440
         Width           =   6492
         _ExtentX        =   11451
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   1
         Min             =   1
         Scrolling       =   1
      End
      Begin VB.Label lblInstallStatus 
         Caption         =   "Status: Copying files"
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2532
      End
      Begin VB.Label Label6 
         Caption         =   "Please wait while the setup installs RAD Basic. This may take several minutes."
         Height          =   372
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   7932
      End
   End
   Begin VB.Frame FrmReadyToInstall 
      BorderStyle     =   0  'None
      Height          =   3372
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   8172
      Begin VB.Label Label5 
         Caption         =   "- Click Cancel to exit"
         Height          =   252
         Left            =   360
         TabIndex        =   22
         Top             =   1680
         Width           =   7452
      End
      Begin VB.Label Label4 
         Caption         =   "- Click Back to change settings"
         Height          =   252
         Left            =   360
         TabIndex        =   21
         Top             =   1320
         Width           =   7452
      End
      Begin VB.Label Label3 
         Caption         =   "- Click Next to begin configuration"
         Height          =   252
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   7212
      End
      Begin VB.Label Label1 
         Caption         =   "The setup is now ready to configure RAD Basic on this computer."
         Height          =   372
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   7572
      End
      Begin VB.Label Label2 
         Caption         =   $"FrmIniSetup.frx":030A
         Height          =   492
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   6372
      End
   End
   Begin VB.Frame FrameFirst 
      BorderStyle     =   0  'None
      Height          =   3132
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   8052
      Begin VB.Label LblNightlyNotice 
         Caption         =   "This installation process only support nightly releases."
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5052
      End
   End
   Begin VB.Frame FrameSelectComponents 
      BorderStyle     =   0  'None
      Height          =   3012
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   8172
      Begin MSComctlLib.TreeView TreeComponents 
         Height          =   2412
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   7692
         _ExtentX        =   13568
         _ExtentY        =   4255
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.Frame FrameEula 
      BorderStyle     =   0  'None
      Height          =   3252
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   8172
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2772
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   7812
         _ExtentX        =   13780
         _ExtentY        =   4890
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         FileName        =   "C:\Users\koss\radbasic-src\rb-installer\EULA.rtf"
         TextRTF         =   $"FrmIniSetup.frx":0361
      End
      Begin VB.CheckBox CKAcceptEula 
         Caption         =   "I accept the terms of license agreement"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   7572
      End
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "< Back"
      Height          =   492
      Left            =   5760
      TabIndex        =   2
      Top             =   4080
      Width           =   1092
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next >"
      Height          =   492
      Left            =   7080
      TabIndex        =   1
      Top             =   4080
      Width           =   1092
   End
   Begin VB.Label lblTitle 
      Caption         =   "RAD Basic Installer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5292
   End
End
Attribute VB_Name = "FrmIniSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' RAD Basic Installer
' Copyright (c) 2019-2025 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.
Option Explicit
Private ReqUninstallOldVer As Boolean
Private IsInstalledAlready As Boolean

Private currentFrame As Frame

Private Enum SetupStep
 IniStep = 0
 ActionStep = 1
 EulaStep = 2
 ComponentStep = 3
 DstFolderStep = 4
 ReadyToInstall = 5
 Installling = 6
End Enum

Private shlShell As Shell32.Shell
Private shlFolder As Shell32.Folder
Private Const BIF_RETURNONLYFSDIRS = &H1

Private CurrenStep As SetupStep


Private Sub CKAcceptEula_Click()

    ' Enable CmdNext button with checked State of EULA
    CmdNext.Enabled = CKAcceptEula.Value
    
End Sub

Private Sub CheckRequirements()
    
    FrmLongOpDialog.SetMessage "Check requirements", "Please, wait while checking requirements..."
    FrmLongOpDialog.Show vbNormal, Me
    
    'MsgBox "test"
    
End Sub

Private Sub CmdChangePath_Click()
    If shlShell Is Nothing Then
        Set shlShell = New Shell32.Shell
    End If
    Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, "Select a Directory", BIF_RETURNONLYFSDIRS)
    If Not shlFolder Is Nothing Then
        TxtDstFolder.Text = shlFolder.Self.Path
    End If
End Sub

Private Sub CmdInstallNew_Click()
    CurrenStep = CurrenStep + 1
    
    ChangeComponent CurrenStep
    UpdateButtons CurrenStep
End Sub

Private Sub CmdBack_Click()
    CurrenStep = CurrenStep - 1
    
    ChangeComponent CurrenStep
    UpdateButtons CurrenStep
End Sub

Private Sub CmdNext_Click()
    Dim installedOk As Boolean

    CurrenStep = CurrenStep + 1
    
    ChangeComponent CurrenStep
    UpdateButtons CurrenStep
    
    If CurrenStep = Installling Then
        ' Force the re
        DoEvents
        
        ' Execute the install
        installedOk = InstallNightly(TxtDstFolder.Text)
        
        If (installedOk) Then
            MsgBox "Installed successfully", vbOKOnly + vbInformation, "RAD Basic Installer"
            ' Refresh the value
            IsInstalledAlready = ReqValidator.IsNewRADBasicInstalled
            ' Refresh Versions step
            PrepareVersionsStep
        Else
            MsgBox "Some error ocurred during the setup process", vbOKOnly + vbExclamation, "RAD Basic Installer"
        End If
        
        ' Back to Main step
        CurrenStep = ActionStep
        ChangeComponent CurrenStep
        UpdateButtons CurrenStep
    End If
    

End Sub

Private Sub ChangeComponent(CurrentStep As SetupStep)
    If Not (currentFrame Is Nothing) Then
        currentFrame.Visible = False
    End If
    
    Select Case CurrenStep
        Case SetupStep.ActionStep
            Set currentFrame = Me.FrameIni
            lblTitle.Caption = "RAD Basic installed versions"
        Case SetupStep.EulaStep
            Set currentFrame = Me.FrameEula
            lblTitle.Caption = "License agreement"
        Case SetupStep.ComponentStep
            Set currentFrame = Me.FrameSelectComponents
            lblTitle.Caption = "Individual components to install"
        Case SetupStep.DstFolderStep
            Set currentFrame = Me.FrmDstFolder
            lblTitle.Caption = "Destination folder"
        Case SetupStep.ReadyToInstall
            Set currentFrame = Me.FrmReadyToInstall
            lblTitle.Caption = "Completing the setup for RAD Basic"
        Case SetupStep.Installling
            Set currentFrame = Me.FrmInstallng
            lblTitle.Caption = "Installing RAD Basic..."
    End Select
    
    currentFrame.Visible = True
End Sub

Private Sub UpdateButtons(CurrentStep As SetupStep)
    Select Case CurrenStep
        Case SetupStep.IniStep
            CmdBack.Enabled = False
            CmdNext.Enabled = True
            
            CmdBack.Visible = False
            CmdNext.Visible = True
        Case SetupStep.ActionStep
            CmdBack.Enabled = False
            CmdNext.Enabled = False
            
            CmdBack.Visible = False
            CmdNext.Visible = False
        Case SetupStep.EulaStep
            CmdBack.Enabled = False
            CmdNext.Enabled = CKAcceptEula.Value ' Special case
            
            CmdBack.Visible = True
            CmdNext.Visible = True
        Case SetupStep.ComponentStep
            CmdBack.Enabled = True
            CmdNext.Enabled = True
            
            CmdBack.Visible = True
            CmdNext.Visible = True
        Case SetupStep.DstFolderStep
            CmdBack.Enabled = True
            CmdNext.Enabled = True
            
            CmdBack.Visible = True
            CmdNext.Visible = True
        Case SetupStep.Installling
            CmdBack.Enabled = False
            CmdNext.Enabled = False
            
            CmdBack.Visible = True
            CmdNext.Visible = True
    End Select

End Sub

Private Sub Form_Load()
    ReqUninstallOldVer = ReqValidator.IsOldRADBasicInstalled
    IsInstalledAlready = ReqValidator.IsNewRADBasicInstalled
        
    ' Log startup
    LogInfo Me.name, "Init RAD Basic Installer. Version: " & App.Major & "." & App.Minor & "." & App.Revision
    LogInfo Me.name, "--------------------------------------------------------"
    LogInfo Me.name, "Param info: "
    LogInfo Me.name, "  IsInstalledAlready: " & IsInstalledAlready
    LogInfo Me.name, "  IsOldRADBasicInstalled: " & IsOldRADBasicInstalled
    ReqUninstallOldVer = ReqValidator.IsOldRADBasicInstalled
    LogInfo Me.name, "--------------------------------------------------------"
    
    ' Prepare Versions Step
    PrepareVersionsStep
    
    ' Populate nodes in treeview (RAD Basic components)
    PopulateTVComponents
    
    ' Default value for destination folder
    TxtDstFolder.Text = Environ$("ProgramFiles(x86)") & "\" & "RAD Basic"
    
    ' Set startup frame/step
    Set currentFrame = FrameFirst
    UpdateButtons CurrenStep
    
End Sub
Public Sub PrepareVersionsStep()
    If IsInstalledAlready Then
        LblNoVersions.Visible = False
        CmdInstallNew.Visible = False
        
        FrmVersion.Visible = True
        FrmVersion.Top = LblNoVersions.Top
    Else
        FrmVersion.Visible = False
    
        LblNoVersions.Visible = True
        CmdInstallNew.Visible = True
    End If
End Sub
Public Sub PopulateTVComponents()
    Dim nodX As Node
    
    ' RAD Basic component (core) is mandatory
    Set nodX = TreeComponents.Nodes.Add(, , "RB", "RAD Basic")
    nodX.Expanded = True
    nodX.Checked = True
    
    ' LLVM backend is mandatory
    Set nodX = TreeComponents.Nodes.Add(, , "LLVM", "LLVM Backend")
    nodX.Expanded = True
    nodX.Checked = True

    ' Additional LLVM backends are optional (only for debug purposes)
    Set nodX = TreeComponents.Nodes.Add(, , "LLVMExtra", "Additional LLVM backends")
    nodX.Expanded = True
    nodX.Checked = False
    
    ' LLVM Debug Symbols are optional (only for debug purposes)
    Set nodX = TreeComponents.Nodes.Add("LLVMExtra", tvwChild, "LLVMExtraDebug", "LLVM Debug Symbols")
    nodX.Expanded = True
    nodX.Checked = False

End Sub

Private Sub TreeComponents_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Expanded = True
End Sub

Private Sub TreeComponents_NodeCheck(ByVal Node As MSComctlLib.Node)

    If Not Node.Parent Is Nothing Then
        If Not Node.Parent.Checked Then
            Node.Parent.Checked = True
        End If
    End If
End Sub


Public Sub SetStep(Message As String, percentVal As Integer)
    Me.lblInstallStatus.Caption = "Status: " & Message
    Me.PgbSetup.Value = percentVal
    
    ' Call to DoEvents to force the repaint
    DoEvents
End Sub

