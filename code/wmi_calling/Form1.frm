VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WMI"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Manaul"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   5895
      Left            =   1260
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   540
      Width           =   8535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   120
      Width           =   8535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear List"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
    'prcDisplayInfo
    prcD
End Sub

Sub prcD()
    Dim WMI_HW
    Set WMI_HW = New WMI_Ascii.WMI_Functions
    
    Text1.Text = Replace(WMI_HW.funWMIHardwareRequest(Trim(Combo1.Text)), ";", vbNewLine)
End Sub


''''''''''''''''''''''''''start testing area'''''''''''''''''''''''''''''''''''
'Win32_Account
'Win32_SystemUsers
Public Function funManualWMI() As String
    Dim sMessageBuilt As String
    Dim strComputer
    Dim objWMIService
    Dim objItem As Object
    Dim colItems
    Dim Win32Name As String
    
    Win32Name = "Win32_ComputerSystem"
    
On Error GoTo ErrWMI:
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from " & Win32Name, , 48)
    sMessageBuilt = Win32Name & vbNewLine
    
On Error Resume Next

    For Each objItem In colItems
    
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ComputerName: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "UserName: " & objItem.UserName & ";"
        sMessageBuilt = sMessageBuilt & "Domain: " & objItem.Domain & ";"
        sMessageBuilt = sMessageBuilt & "PrimaryOwnerName: " & objItem.PrimaryOwnerName & ";"
        sMessageBuilt = sMessageBuilt & "SystemType: " & objItem.SystemType & ";"
        sMessageBuilt = sMessageBuilt & "BootupState: " & objItem.BootupState & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model & ";"
        sMessageBuilt = sMessageBuilt & "NumberOfProcessors: " & objItem.NumberOfProcessors & ";"
        sMessageBuilt = sMessageBuilt & "OEMLogoBitmap: " & objItem.OEMLogoBitmap & ";"
        sMessageBuilt = sMessageBuilt & "PartOfDomain: " & objItem.PartOfDomain & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "ThermalState: " & objItem.ThermalState & ";"
        sMessageBuilt = sMessageBuilt & "SystemStartupDelay: " & objItem.SystemStartupDelay & ";"
        sMessageBuilt = sMessageBuilt & "SystemStartupSetting: " & objItem.SystemStartupSetting & ";"
        sMessageBuilt = sMessageBuilt & "TotalPhysicalMemory: " & objItem.TotalPhysicalMemory & ";"
        sMessageBuilt = sMessageBuilt & "Workgroup: " & objItem.Workgroup & ";"
        sMessageBuilt = sMessageBuilt & "OEMStringArray: " & objItem.OEMStringArray & ";"
        sMessageBuilt = sMessageBuilt & "Roles: " & objItem.Roles & ";"
        sMessageBuilt = sMessageBuilt & "SystemStartupOptions: " & objItem.SystemStartupOptions & ";"
        
    Next
            
            
    funManualWMI = sMessageBuilt
    Exit Function
    
ErrWMI:
    funManualWMI = "Err (" & Err.Number & ") " & Err.Description & vbNewLine & _
                    "Message Recovered: " & sMessageBuilt
    Exit Function
End Function

Public Function funManualWMI2() As String
    Dim sMessageBuilt As String
    Dim strComputer
    Dim objWMIService
    Dim objItem
    Dim colItems
    Dim System
    
On Error GoTo ErrWMI:
    
    Set System = GetObject("winmgmts:{impersonationLevel=" & _
                        "impersonate}!\\modcon002\root\cimv2:" & _
                        "Win32_ComputerSystem=""modcon002""")
    
    
    sMessageBuilt = sMessageBuilt & System.Caption & vbNewLine
    sMessageBuilt = sMessageBuilt & System.Domain & vbNewLine
    sMessageBuilt = sMessageBuilt & "PrimaryOwnerName: " & objItem.PrimaryOwnerName & vbNewLine
    sMessageBuilt = sMessageBuilt & "SystemType: " & objItem.SystemType & vbNewLine
            
            
    funManualWMI2 = sMessageBuilt
    Exit Function
    
ErrWMI:
    sMessageBuilt = sMessageBuilt & "Err (" & Err.Number & ") " & Err.Description & vbNewLine
    'funManualWMI2 = "Err (" & Err.Number & ") " & Err.Description & vbNewLine & _
                    '"Message Recovered: " & sMessageBuilt
    Resume Next
    Exit Function
End Function

Private Sub Command3_Click()
    
    Text1.Text = Replace(funManualWMI, ";", vbNewLine)
    'Text1.Text = funManualWMI2
End Sub

''''''''''''''''''''''''''''''end testing'''''''''''''''''''''''''''''''''''''

Sub prcFillCombo()
    Combo1.Text = "prcWin32_DiskDrive"
    Combo1.AddItem "prcWin32_DiskDrive"
    Combo1.AddItem "Win32_DisplayConfiguration"
    Combo1.AddItem "Win32_DisplayControllerConfiguration"
    Combo1.AddItem "Win32_Processor"
    Combo1.AddItem "Win32_SMBIOSMemory"
    Combo1.AddItem "Win32_SoundDevice"
    Combo1.AddItem "Win32_SystemEnclosure"
    Combo1.AddItem "Win32_SystemMemoryResource"
    Combo1.AddItem "Win32_SystemSlot"
    Combo1.AddItem "Win32_TemperatureProbe"
    Combo1.AddItem "Win32_USBController"
    Combo1.AddItem "Win32_VideoConfiguration"
    Combo1.AddItem "Win32_VideoController"
    Combo1.AddItem "Win32_VoltageProbe"
    Combo1.AddItem "Win32_1394Controller"
    Combo1.AddItem "Win32_BaseBoard"
    Combo1.AddItem "Win32_Battery"
    Combo1.AddItem "Win32_BIOS"
    Combo1.AddItem "Win32_Bus"
    Combo1.AddItem "Win32_CacheMemory"
    Combo1.AddItem "Win32_CDROMDrive"
    Combo1.AddItem "Win32_CurrentProbe"
    Combo1.AddItem "Win32_DesktopMonitor"
    Combo1.AddItem "Win32_DeviceMemoryAddress"
    Combo1.AddItem "Win32_DMAChannel"
    Combo1.AddItem "Win32_Fan"
    Combo1.AddItem "Win32_FloppyController"
    Combo1.AddItem "Win32_FloppyDrive"
    Combo1.AddItem "Win32_IDEController"
    Combo1.AddItem "Win32_IRQResource"
    Combo1.AddItem "Win32_Keyboard"
    Combo1.AddItem "Win32_MemoryArray"
    Combo1.AddItem "Win32_MemoryDevice"
    Combo1.AddItem "Win32_MotherboardDevice"
    Combo1.AddItem "Win32_NetworkAdapter"
    Combo1.AddItem "Win32_NetworkAdapterConfiguration"
    Combo1.AddItem "Win32_OnBoardDevice"
    Combo1.AddItem "Win32_ParallelPort"
    Combo1.AddItem "Win32_PhysicalMemory"
    Combo1.AddItem "Win32_PhysicalMemoryArray"
    Combo1.AddItem "Win32_PnPEntity"
    Combo1.AddItem "Win32_PointingDevice"
    Combo1.AddItem "Win32_PortConnector"
    Combo1.AddItem "Win32_PortResource"
    Combo1.AddItem "Win32_POTSModem"
    Combo1.AddItem "Win32_PowerManagementEvent"
    Combo1.AddItem "Win32_Printer"
    Combo1.AddItem "Win32_PrinterConfiguration"
    Combo1.AddItem "Win32_PrintJob"
    Combo1.AddItem "Win32_SerialPort"
    Combo1.AddItem "Win32_SerialPortConfiguration"
    Combo1.AddItem "----------------------------"
    Combo1.AddItem "Win32_QuotaSetting"
    Combo1.AddItem "Win32_OperatingSystem"
    Combo1.AddItem "Win32_ComputerSystem"
    
End Sub


Private Sub Form_Load()
    prcFillCombo
End Sub
