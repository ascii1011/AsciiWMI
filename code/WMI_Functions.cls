VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WMI_Functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'''''''''''''''''''''''''''''''''''
'call by this method
'
'        Dim WMI_Hardware
'
'        Private Sub Command1_Click()
'            Dim stemp As String
'
'            stemp = WMI_Hardware.funWMIHardwareRequest(Trim(cboName.Text))
'
'            MsgBox stemp
'        End Sub
'
'        Private Sub Form_Load()
'            prcInitVars
'        End Sub
'
'        Sub prcInitVars()
'            Dim l As Long
'            Set WMI_Hardware = New WMI_Ascii.WMI_Functions
'
'            For l = 0 To WMI_Hardware.WMIHardwareItemCount
'                If l = 0 Then cboName.Text = WMI_Hardware.funRequestSpecificWMIHardwareItem(l)
'                cboName.AddItem WMI_Hardware.funRequestSpecificWMIHardwareItem(l)
'            Next l
'        End Sub
'
''''''''''''''''''''''''''''''''''''


Private sMessageBuilt As String
Private sWMIHardwareItems() As String

Public Function Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Function


Public Function PopulateCombo() As String
    Dim i As Integer
    prcInitWMIHardwareItems
    For i = 0 To UBound(sWMIHardwareItems) - 1
        PopulateCombo = PopulateCombo & sWMIHardwareItems(i) & ";"
    Next i
End Function

Public Function PopulateCombo_Long() As String
    Exit Function
    cboName.Text = "prcWin32_DiskDrive"
    cboName.AddItem "prcWin32_DiskDrive"
    cboName.AddItem "Win32_DisplayConfiguration"
    cboName.AddItem "Win32_DisplayControllerConfiguration"
    cboName.AddItem "Win32_Processor"
    cboName.AddItem "Win32_SMBIOSMemory"
    cboName.AddItem "Win32_SoundDevice"
    cboName.AddItem "Win32_SystemEnclosure"
    cboName.AddItem "Win32_SystemMemoryResource"
    cboName.AddItem "Win32_SystemSlot"
    cboName.AddItem "Win32_TemperatureProbe"
    cboName.AddItem "Win32_USBController"
    cboName.AddItem "Win32_VideoConfiguration"
    cboName.AddItem "Win32_VideoController"
    cboName.AddItem "Win32_VoltageProbe"
    cboName.AddItem "Win32_1394Controller"
    cboName.AddItem "Win32_BaseBoard"
    cboName.AddItem "Win32_Battery"
    cboName.AddItem "Win32_BIOS"
    cboName.AddItem "Win32_Bus"
    cboName.AddItem "Win32_CacheMemory"
    cboName.AddItem "Win32_CDROMDrive"
    cboName.AddItem "Win32_CurrentProbe"
    cboName.AddItem "Win32_DesktopMonitor"
    cboName.AddItem "Win32_DeviceMemoryAddress"
    cboName.AddItem "Win32_DMAChannel"
    cboName.AddItem "Win32_Fan"
    cboName.AddItem "Win32_FloppyController"
    cboName.AddItem "Win32_FloppyDrive"
    cboName.AddItem "Win32_IDEController"
    cboName.AddItem "Win32_IRQResource"
    cboName.AddItem "Win32_Keyboard"
    cboName.AddItem "Win32_MemoryArray"
    cboName.AddItem "Win32_MemoryDevice"
    cboName.AddItem "Win32_MotherboardDevice"
    cboName.AddItem "Win32_NetworkAdapter"
    cboName.AddItem "Win32_NetworkAdapterConfiguration"
    cboName.AddItem "Win32_OnBoardDevice"
    cboName.AddItem "Win32_ParallelPort"
    cboName.AddItem "Win32_PhysicalMemory"
    cboName.AddItem "Win32_PhysicalMemoryArray"
    cboName.AddItem "Win32_PnPEntity"
    cboName.AddItem "Win32_PointingDevice"
    cboName.AddItem "Win32_PortConnector"
    cboName.AddItem "Win32_PortResource"
    cboName.AddItem "Win32_POTSModem"
    cboName.AddItem "Win32_PowerManagementEvent"
    cboName.AddItem "Win32_Printer"
    cboName.AddItem "Win32_PrinterConfiguration"
    cboName.AddItem "Win32_PrintJob"
    cboName.AddItem "Win32_SerialPort"
    cboName.AddItem "Win32_SerialPortConfiguration"
    cboName.AddItem "Win32_ComputerSystem"
    cboName.AddItem "Win32_QuotaSetting"
    cboName.AddItem "Win32_OperatingSystem"
    
End Function

Public Function funRequestSpecificWMIHardwareItem(lIndex As Long) As String
    '0 to ubound -1
On Error GoTo RequestErr:
    
    If UBound(sWMIHardwareItems) >= lIndex And 0 <= lIndex Then
        funRequestSpecificWMIHardwareItem = sWMIHardwareItems(lIndex)
    Else
        funRequestSpecificWMIHardwareItem = UBound(sWMIHardwareItems) & ":" & lIndex & ":Error Index!"
    End If
    Exit Function
    
RequestErr:
    funRequestSpecificWMIHardwareItem = "Err: " & Err.Number & ":" & Err.Description
End Function

Public Function WMIHardwareItemCount() As Long
    prcInitWMIHardwareItems
    WMIHardwareItemCount = UBound(sWMIHardwareItems)
End Function

Private Sub prcInitWMIHardwareItems()

    ReDim sWMIHardwareItems(52)
    
    sWMIHardwareItems(0) = "Win32_DiskDrive"
    sWMIHardwareItems(1) = "Win32_DisplayConfiguration"
    sWMIHardwareItems(2) = "Win32_DisplayControllerConfiguration"
    sWMIHardwareItems(3) = "Win32_Processor"
    sWMIHardwareItems(4) = "Win32_SMBIOSMemory"
    sWMIHardwareItems(5) = "Win32_SoundDevice"
    sWMIHardwareItems(6) = "Win32_SystemEnclosure"
    sWMIHardwareItems(7) = "Win32_SystemMemoryResource"
    sWMIHardwareItems(8) = "Win32_SystemSlot"
    sWMIHardwareItems(9) = "Win32_TemperatureProbe"
    sWMIHardwareItems(10) = "Win32_USBController"
    sWMIHardwareItems(11) = "Win32_VideoConfiguration"
    sWMIHardwareItems(12) = "Win32_VideoController"
    sWMIHardwareItems(13) = "Win32_VoltageProbe"
    sWMIHardwareItems(14) = "Win32_1394Controller"
    sWMIHardwareItems(15) = "Win32_BaseBoard"
    sWMIHardwareItems(16) = "Win32_Battery"
    sWMIHardwareItems(17) = "Win32_BIOS"
    sWMIHardwareItems(18) = "Win32_Bus"
    sWMIHardwareItems(19) = "Win32_CacheMemory"
    sWMIHardwareItems(20) = "Win32_CDROMDrive"
    sWMIHardwareItems(21) = "Win32_CurrentProbe"
    sWMIHardwareItems(22) = "Win32_DesktopMonitor"
    sWMIHardwareItems(23) = "Win32_DeviceMemoryAddress"
    sWMIHardwareItems(24) = "Win32_DMAChannel"
    sWMIHardwareItems(25) = "Win32_Fan"
    sWMIHardwareItems(26) = "Win32_FloppyController"
    sWMIHardwareItems(27) = "Win32_FloppyDrive"
    sWMIHardwareItems(28) = "Win32_IDEController"
    sWMIHardwareItems(29) = "Win32_IRQResource"
    sWMIHardwareItems(30) = "Win32_Keyboard"
    sWMIHardwareItems(31) = "Win32_MemoryArray"
    sWMIHardwareItems(32) = "Win32_MemoryDevice"
    sWMIHardwareItems(33) = "Win32_MotherboardDevice"
    sWMIHardwareItems(34) = "Win32_NetworkAdapter"
    sWMIHardwareItems(35) = "Win32_NetworkAdapterConfiguration"
    sWMIHardwareItems(36) = "Win32_OnBoardDevice"
    sWMIHardwareItems(37) = "Win32_ParallelPort"
    sWMIHardwareItems(38) = "Win32_PhysicalMemory"
    sWMIHardwareItems(39) = "Win32_PhysicalMemoryArray"
    sWMIHardwareItems(40) = "Win32_PnPEntity"
    sWMIHardwareItems(41) = "Win32_PointingDevice"
    sWMIHardwareItems(42) = "Win32_PortConnector"
    sWMIHardwareItems(43) = "Win32_PortResource"
    sWMIHardwareItems(44) = "Win32_POTSModem"
    sWMIHardwareItems(45) = "Win32_PowerManagementEvent"
    sWMIHardwareItems(46) = "Win32_Printer"
    sWMIHardwareItems(47) = "Win32_PrinterConfiguration"
    sWMIHardwareItems(48) = "Win32_PrintJob"
    sWMIHardwareItems(49) = "Win32_SerialPort"
    sWMIHardwareItems(50) = "Win32_SerialPortConfiguration"
    sWMIHardwareItems(51) = "Win32_Service"
    sWMIHardwareItems(52) = "Win32_QuotaSetting"
    sWMIHardwareItems(53) = "Win32_OperatingSystem"
    sWMIHardwareItems(54) = "Win32_ComputerSystem"
    
End Sub



Public Function funWMIHardwareRequest(sName) As String
    Dim strComputer
    Dim objWMIService
    Dim objItem
    Dim colItems
    
On Error GoTo ErrWMI:
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from " & sName, , 48)
    sMessageBuilt = sName & vbNewLine
    
On Error Resume Next

    For Each objItem In colItems
    
        Select Case sName
        
            Case "Win32_DiskDrive"
                    Win32_DiskDrive objItem
            Case "Win32_DisplayConfiguration"
                    Win32_DisplayConfiguration objItem
            Case "Win32_DisplayControllerConfiguration"
                    Win32_DisplayControllerConfiguration objItem
            Case "Win32_Processor"
                    Win32_Processor objItem
            Case "Win32_SMBIOSMemory"
                    Win32_SMBIOSMemory objItem
            Case "Win32_SoundDevice"
                    Win32_SoundDevice objItem
            Case "Win32_SystemEnclosure"
                    Win32_SystemEnclosure objItem
            Case "Win32_SystemMemoryResource"
                    Win32_SystemMemoryResource objItem
            Case "Win32_SystemSlot"
                    Win32_SystemSlot objItem
            Case "Win32_TemperatureProbe"
                    Win32_TemperatureProbe objItem
            Case "Win32_USBController"
                    Win32_USBController objItem
            Case "Win32_VideoConfiguration"
                    Win32_VideoConfiguration objItem
            Case "Win32_VideoController"
                    Win32_VideoController objItem
            Case "Win32_VoltageProbe"
                    Win32_VoltageProbe objItem
            Case "Win32_1394Controller"
                    Win32_1394Controller objItem
            Case "Win32_BaseBoard"
                    Win32_BaseBoard objItem
            Case "Win32_Battery"
                    Win32_Battery objItem
            Case "Win32_BIOS"
                    Win32_BIOS objItem
            Case "Win32_Bus"
                    Win32_Bus objItem
            Case "Win32_CacheMemory"
                    Win32_CacheMemory objItem
            Case "Win32_CDROMDrive"
                    Win32_CDROMDrive objItem
            Case "Win32_CurrentProbe"
                    Win32_CurrentProbe objItem
            Case "Win32_DesktopMonitor"
                    Win32_DesktopMonitor objItem
            Case "Win32_DeviceMemoryAddress"
                    Win32_DeviceMemoryAddress objItem
            Case "Win32_DMAChannel"
                    Win32_DMAChannel objItem
            Case "Win32_Fan"
                    Win32_Fan objItem
            Case "Win32_FloppyController"
                    Win32_FloppyController objItem
            Case "Win32_FloppyDrive"
                    Win32_FloppyDrive objItem
            Case "Win32_IDEController"
                    Win32_IDEController objItem
            Case "Win32_IRQResource"
                    Win32_IRQResource objItem
            Case "Win32_Keyboard"
                    Win32_Keyboard objItem
            Case "Win32_MemoryArray"
                    Win32_MemoryArray objItem
            Case "Win32_MemoryDevice"
                    Win32_MemoryDevice objItem
            Case "Win32_MotherboardDevice"
                    Win32_MotherboardDevice objItem
            Case "Win32_NetworkAdapter"
                    Win32_NetworkAdapter objItem
            Case "Win32_NetworkAdapterConfiguration"
                    Win32_NetworkAdapterConfiguration objItem
            Case "Win32_OnBoardDevice"
                    Win32_OnBoardDevice objItem
            Case "Win32_ParallelPort"
                    Win32_ParallelPort objItem
            Case "Win32_PhysicalMemory"
                    Win32_PhysicalMemory objItem
            Case "Win32_PhysicalMemoryArray"
                    Win32_PhysicalMemoryArray objItem
            Case "Win32_PnPEntity"
                    Win32_PnPEntity objItem
            Case "Win32_PointingDevice"
                    Win32_PointingDevice objItem
            Case "Win32_PortConnector"
                    Win32_PortConnector objItem
            Case "Win32_PortResource"
                    Win32_PortResource objItem
            Case "Win32_POTSModem"
                    Win32_POTSModem objItem
            Case "Win32_PowerManagementEvent"
                    Win32_PowerManagementEvent objItem
            Case "Win32_Printer"
                    Win32_Printer objItem
            Case "Win32_PrinterConfiguration"
                    Win32_PrinterConfiguration objItem
            Case "Win32_PrintJob"
                    Win32_PrintJob objItem
            Case "Win32_SerialPort"
                    Win32_SerialPort objItem
            Case "Win32_SerialPortConfiguration"
                    Win32_SerialPortConfiguration objItem
            Case "Win32_Service"
                    Win32_Service objItem
            Case "Win32_QuotaSetting"
                    Win32_QuotaSetting objItem
            Case "Win32_OperatingSystem"
                    Win32_OperatingSystem objItem
            Case "Win32_ComputerSystem"
                    Win32_ComputerSystem objItem
            
        End Select
        
    Next
            
            
    funWMIHardwareRequest = sMessageBuilt
    Exit Function
    
ErrWMI:
    funWMIHardwareRequest = "Err (" & Err.Number & ") " & Err.Description & vbNewLine & _
                    "Message Recovered: " & sMessageBuilt
    Exit Function
            
            
End Function

Public Sub Win32_ComputerSystem(objItem)

On Error Resume Next

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
        sMessageBuilt = sMessageBuilt & "OEMStringArray: " & objItem.OEMStringArray & ";"
        sMessageBuilt = sMessageBuilt & "Roles: " & objItem.Roles & ";"
        sMessageBuilt = sMessageBuilt & "PartOfDomain: " & objItem.PartOfDomain & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "ThermalState: " & objItem.ThermalState & ";"
        sMessageBuilt = sMessageBuilt & "SystemStartupDelay: " & objItem.SystemStartupDelay & ";"
        sMessageBuilt = sMessageBuilt & "SystemStartupOptions: " & objItem.SystemStartupOptions & ";"
        sMessageBuilt = sMessageBuilt & "SystemStartupSetting: " & objItem.SystemStartupSetting & ";"
        sMessageBuilt = sMessageBuilt & "TotalPhysicalMemory: " & objItem.TotalPhysicalMemory & ";"
        sMessageBuilt = sMessageBuilt & "Workgroup: " & objItem.Workgroup & ";"
        
End Sub

Public Sub Win32_QuotaSetting(objItem)

On Error Resume Next

        sMessageBuilt = sMessageBuilt & "BitsPerPel: " & objItem.BitsPerPel & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceName: " & objItem.DeviceName & ";"
        sMessageBuilt = sMessageBuilt & "DisplayFlags: " & objItem.DisplayFlags & ";"
        sMessageBuilt = sMessageBuilt & "DisplayFrequency: " & objItem.DisplayFrequency & ";"
        sMessageBuilt = sMessageBuilt & "DitherType: " & objItem.DitherType & ";"
        sMessageBuilt = sMessageBuilt & "DriverVersion: " & objItem.DriverVersion & ";"
        sMessageBuilt = sMessageBuilt & "ICMIntent: " & objItem.ICMIntent & ";"
        sMessageBuilt = sMessageBuilt & "ICMMethod: " & objItem.ICMMethod & ";"
        sMessageBuilt = sMessageBuilt & "LogPixels: " & objItem.LogPixels & ";"
        sMessageBuilt = sMessageBuilt & "PelsHeight: " & objItem.PelsHeight & ";"
        sMessageBuilt = sMessageBuilt & "PelsWidth: " & objItem.PelsWidth & ";"
        sMessageBuilt = sMessageBuilt & "SettingID: " & objItem.SettingID & ";"
        sMessageBuilt = sMessageBuilt & "SpecificationVersion: " & objItem.SpecificationVersion
        
End Sub
Public Sub Win32_OperatingSystem(objItem)

On Error Resume Next

        sMessageBuilt = sMessageBuilt & "BitsPerPel: " & objItem.BitsPerPel & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceName: " & objItem.DeviceName & ";"
        sMessageBuilt = sMessageBuilt & "DisplayFlags: " & objItem.DisplayFlags & ";"
        sMessageBuilt = sMessageBuilt & "DisplayFrequency: " & objItem.DisplayFrequency & ";"
        sMessageBuilt = sMessageBuilt & "DitherType: " & objItem.DitherType & ";"
        sMessageBuilt = sMessageBuilt & "DriverVersion: " & objItem.DriverVersion & ";"
        sMessageBuilt = sMessageBuilt & "ICMIntent: " & objItem.ICMIntent & ";"
        sMessageBuilt = sMessageBuilt & "ICMMethod: " & objItem.ICMMethod & ";"
        sMessageBuilt = sMessageBuilt & "LogPixels: " & objItem.LogPixels & ";"
        sMessageBuilt = sMessageBuilt & "PelsHeight: " & objItem.PelsHeight & ";"
        sMessageBuilt = sMessageBuilt & "PelsWidth: " & objItem.PelsWidth & ";"
        sMessageBuilt = sMessageBuilt & "SettingID: " & objItem.SettingID & ";"
        sMessageBuilt = sMessageBuilt & "SpecificationVersion: " & objItem.SpecificationVersion
        
End Sub

Public Sub Win32_DisplayConfiguration(objItem)

On Error Resume Next

        sMessageBuilt = sMessageBuilt & "BitsPerPel: " & objItem.BitsPerPel & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceName: " & objItem.DeviceName & ";"
        sMessageBuilt = sMessageBuilt & "DisplayFlags: " & objItem.DisplayFlags & ";"
        sMessageBuilt = sMessageBuilt & "DisplayFrequency: " & objItem.DisplayFrequency & ";"
        sMessageBuilt = sMessageBuilt & "DitherType: " & objItem.DitherType & ";"
        sMessageBuilt = sMessageBuilt & "DriverVersion: " & objItem.DriverVersion & ";"
        sMessageBuilt = sMessageBuilt & "ICMIntent: " & objItem.ICMIntent & ";"
        sMessageBuilt = sMessageBuilt & "ICMMethod: " & objItem.ICMMethod & ";"
        sMessageBuilt = sMessageBuilt & "LogPixels: " & objItem.LogPixels & ";"
        sMessageBuilt = sMessageBuilt & "PelsHeight: " & objItem.PelsHeight & ";"
        sMessageBuilt = sMessageBuilt & "PelsWidth: " & objItem.PelsWidth & ";"
        sMessageBuilt = sMessageBuilt & "SettingID: " & objItem.SettingID & ";"
        sMessageBuilt = sMessageBuilt & "SpecificationVersion: " & objItem.SpecificationVersion
        
End Sub





Sub Win32_DiskDrive(objItem)
    

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "BytesPerSector: " & objItem.BytesPerSector & ";"
        sMessageBuilt = sMessageBuilt & "Capabilities: " & objItem.Capabilities & ";"
        sMessageBuilt = sMessageBuilt & "CapabilityDescriptions: " & objItem.CapabilityDescriptions & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "CompressionMethod: " & objItem.CompressionMethod & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "DefaultBlockSize: " & objItem.DefaultBlockSize & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "ErrorMethodology: " & objItem.ErrorMethodology & ";"
        sMessageBuilt = sMessageBuilt & "Index: " & objItem.Index & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "InterfaceType: " & objItem.InterfaceType & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MaxBlockSize: " & objItem.MaxBlockSize & ";"
        sMessageBuilt = sMessageBuilt & "MaxMediaSize: " & objItem.MaxMediaSize & ";"
        sMessageBuilt = sMessageBuilt & "MediaLoaded: " & objItem.MediaLoaded & ";"
        sMessageBuilt = sMessageBuilt & "MediaType: " & objItem.MediaType & ";"
        sMessageBuilt = sMessageBuilt & "MinBlockSize: " & objItem.MinBlockSize & ";"
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NeedsCleaning: " & objItem.NeedsCleaning & ";"
        sMessageBuilt = sMessageBuilt & "NumberOfMediaSupported: " & objItem.NumberOfMediaSupported & ";"
        sMessageBuilt = sMessageBuilt & "Partitions: " & objItem.Partitions & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "SCSIBus: " & objItem.SCSIBus & ";"
        sMessageBuilt = sMessageBuilt & "SCSILogicalUnit: " & objItem.SCSILogicalUnit & ";"
        sMessageBuilt = sMessageBuilt & "SCSIPort: " & objItem.SCSIPort & ";"
        sMessageBuilt = sMessageBuilt & "SCSITargetId: " & objItem.SCSITargetId & ";"
        sMessageBuilt = sMessageBuilt & "SectorsPerTrack: " & objItem.SectorsPerTrack & ";"
        sMessageBuilt = sMessageBuilt & "Size: " & objItem.Size & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        sMessageBuilt = sMessageBuilt & "TotalCylinders: " & objItem.TotalCylinders & ";"
        sMessageBuilt = sMessageBuilt & "TotalHeads: " & objItem.TotalHeads & ";"
        sMessageBuilt = sMessageBuilt & "TotalSectors: " & objItem.TotalSectors & ";"
        sMessageBuilt = sMessageBuilt & "TotalTracks: " & objItem.TotalTracks & ";"
        sMessageBuilt = sMessageBuilt & "TracksPerCylinder: " & objItem.TracksPerCylinder
        
    
End Sub


Sub Win32_DisplayControllerConfiguration(objItem)
    

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_DisplayControllerConfiguration", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "BitsPerPixel: " & objItem.BitsPerPixel & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ColorPlanes: " & objItem.ColorPlanes & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceEntriesInAColorTable: " & objItem.DeviceEntriesInAColorTable & ";"
        sMessageBuilt = sMessageBuilt & "DeviceSpecificPens: " & objItem.DeviceSpecificPens & ";"
        sMessageBuilt = sMessageBuilt & "HorizontalResolution: " & objItem.HorizontalResolution & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "RefreshRate: " & objItem.RefreshRate & ";"
        sMessageBuilt = sMessageBuilt & "ReservedSystemPaletteEntries: " & objItem.ReservedSystemPaletteEntries & ";"
        sMessageBuilt = sMessageBuilt & "SettingID: " & objItem.SettingID & ";"
        sMessageBuilt = sMessageBuilt & "SystemPaletteEntries: " & objItem.SystemPaletteEntries & ";"
        sMessageBuilt = sMessageBuilt & "VerticalResolution: " & objItem.VerticalResolution & ";"
        sMessageBuilt = sMessageBuilt & "VideoMode: " & objItem.VideoMode
        
    
End Sub


Sub Win32_Processor(objItem)

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "AddressWidth: " & objItem.AddressWidth & ";"
        sMessageBuilt = sMessageBuilt & "Architecture: " & objItem.Architecture & ";"
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CpuStatus: " & objItem.CpuStatus & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "CurrentClockSpeed: " & objItem.CurrentClockSpeed & ";"
        sMessageBuilt = sMessageBuilt & "CurrentVoltage: " & objItem.CurrentVoltage & ";"
        sMessageBuilt = sMessageBuilt & "DataWidth: " & objItem.DataWidth & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "ExtClock: " & objItem.ExtClock & ";"
        sMessageBuilt = sMessageBuilt & "Family: " & objItem.Family & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "L2CacheSize: " & objItem.L2CacheSize & ";"
        sMessageBuilt = sMessageBuilt & "L2CacheSpeed: " & objItem.L2CacheSpeed & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "Level: " & objItem.Level & ";"
        sMessageBuilt = sMessageBuilt & "LoadPercentage: " & objItem.LoadPercentage & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MaxClockSpeed: " & objItem.MaxClockSpeed & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "OtherFamilyDescription: " & objItem.OtherFamilyDescription & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "ProcessorId: " & objItem.ProcessorId & ";"
        sMessageBuilt = sMessageBuilt & "ProcessorType: " & objItem.ProcessorType & ";"
        sMessageBuilt = sMessageBuilt & "Revision: " & objItem.Revision & ";"
        sMessageBuilt = sMessageBuilt & "Role: " & objItem.Role & ";"
        sMessageBuilt = sMessageBuilt & "SocketDesignation: " & objItem.SocketDesignation & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "Stepping: " & objItem.Stepping & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        sMessageBuilt = sMessageBuilt & "UniqueId: " & objItem.UniqueId & ";"
        sMessageBuilt = sMessageBuilt & "UpgradeMethod: " & objItem.UpgradeMethod & ";"
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Version & ";"
        sMessageBuilt = sMessageBuilt & "VoltageCaps: " & objItem.VoltageCaps
        
    
End Sub




Sub Win32_SMBIOSMemory(objItem)
    
    

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SMBIOSMemory", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Access: " & objItem.Access & ";"
        sMessageBuilt = sMessageBuilt & "AdditionalErrorData: " & objItem.AdditionalErrorData & ";"
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "BlockSize: " & objItem.BlockSize & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CorrectableError: " & objItem.CorrectableError & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "EndingAddress: " & objItem.EndingAddress & ";"
        sMessageBuilt = sMessageBuilt & "ErrorAccess: " & objItem.ErrorAccess & ";"
        sMessageBuilt = sMessageBuilt & "ErrorAddress: " & objItem.ErrorAddress & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorData: " & objItem.ErrorData & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDataOrder: " & objItem.ErrorDataOrder & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "ErrorInfo: " & objItem.ErrorInfo & ";"
        sMessageBuilt = sMessageBuilt & "ErrorMethodology: " & objItem.ErrorMethodology & ";"
        sMessageBuilt = sMessageBuilt & "ErrorResolution: " & objItem.ErrorResolution & ";"
        sMessageBuilt = sMessageBuilt & "ErrorTime: " & objItem.ErrorTime & ";"
        sMessageBuilt = sMessageBuilt & "ErrorTransferSize: " & objItem.ErrorTransferSize & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NumberOfBlocks: " & objItem.NumberOfBlocks & ";"
        sMessageBuilt = sMessageBuilt & "OtherErrorDescription: " & objItem.OtherErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "Purpose: " & objItem.Purpose & ";"
        sMessageBuilt = sMessageBuilt & "StartingAddress: " & objItem.StartingAddress & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemLevelAddress: " & objItem.SystemLevelAddress & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        
    
End Sub

Sub Win32_SoundDevice(objItem)
    

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SoundDevice", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "DMABufferSize: " & objItem.DMABufferSize & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MPU401Address: " & objItem.MPU401Address & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "ProductName: " & objItem.ProductName & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        
    
End Sub

Sub Win32_SystemEnclosure(objItem)
    

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "AudibleAlarm: " & objItem.AudibleAlarm & ";"
        sMessageBuilt = sMessageBuilt & "BreachDescription: " & objItem.BreachDescription & ";"
        sMessageBuilt = sMessageBuilt & "CableManagementStrategy: " & objItem.CableManagementStrategy & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ChassisTypes: " & objItem.ChassisTypes & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "CurrentRequiredOrProduced: " & objItem.CurrentRequiredOrProduced & ";"
        sMessageBuilt = sMessageBuilt & "Depth: " & objItem.Depth & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "HeatGeneration: " & objItem.HeatGeneration & ";"
        sMessageBuilt = sMessageBuilt & "Height: " & objItem.Height & ";"
        sMessageBuilt = sMessageBuilt & "HotSwappable: " & objItem.HotSwappable & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "LockPresent: " & objItem.LockPresent & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NumberOfPowerCords: " & objItem.NumberOfPowerCords & ";"
        sMessageBuilt = sMessageBuilt & "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo & ";"
        sMessageBuilt = sMessageBuilt & "PartNumber: " & objItem.PartNumber & ";"
        sMessageBuilt = sMessageBuilt & "PoweredOn: " & objItem.PoweredOn & ";"
        sMessageBuilt = sMessageBuilt & "Removable: " & objItem.Removable & ";"
        sMessageBuilt = sMessageBuilt & "Replaceable: " & objItem.Replaceable & ";"
        sMessageBuilt = sMessageBuilt & "SecurityBreach: " & objItem.SecurityBreach & ";"
        sMessageBuilt = sMessageBuilt & "SecurityStatus: " & objItem.SecurityStatus & ";"
        sMessageBuilt = sMessageBuilt & "SerialNumber: " & objItem.SerialNumber & ";"
        sMessageBuilt = sMessageBuilt & "ServiceDescriptions: " & objItem.ServiceDescriptions & ";"
        sMessageBuilt = sMessageBuilt & "ServicePhilosophy: " & objItem.ServicePhilosophy & ";"
        sMessageBuilt = sMessageBuilt & "SKU: " & objItem.SKU & ";"
        sMessageBuilt = sMessageBuilt & "SMBIOSAssetTag: " & objItem.SMBIOSAssetTag & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "Tag: " & objItem.Tag & ";"
        sMessageBuilt = sMessageBuilt & "TypeDescriptions: " & objItem.TypeDescriptions & ";"
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Version & ";"
        sMessageBuilt = sMessageBuilt & "VisibleAlarm: " & objItem.VisibleAlarm & ";"
        sMessageBuilt = sMessageBuilt & "Weight: " & objItem.Weight & ";"
        sMessageBuilt = sMessageBuilt & "Width: " & objItem.Width
        
    
End Sub

Sub Win32_SystemMemoryResource(objItem)
    'a little exhausting
    

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SystemMemoryResource", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "CSCreationClassName: " & objItem.CSCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "CSName: " & objItem.CSName & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "EndingAddress: " & objItem.EndingAddress & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "StartingAddress: " & objItem.StartingAddress & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        
    
End Sub

Sub Win32_SystemSlot(objItem)
    'a little exhausting

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SystemSlot", , 48)
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Captions & ";"
        sMessageBuilt = sMessageBuilt & "ConnectorPinout: " & objItem.ConnectorPinouts & ";"
        sMessageBuilt = sMessageBuilt & "ConnectorType: " & objItem.ConnectorTypes & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassNames & ";"
        sMessageBuilt = sMessageBuilt & "CurrentUsage: " & objItem.CurrentUsages & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Descriptions & ";"
        sMessageBuilt = sMessageBuilt & "HeightAllowed: " & objItem.HeightAlloweds & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDates & ";"
        sMessageBuilt = sMessageBuilt & "LengthAllowed: " & objItem.LengthAlloweds & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturers & ";"
        sMessageBuilt = sMessageBuilt & "MaxDataWidth: " & objItem.MaxDataWidths & ";"
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Models & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Names & ";"
        sMessageBuilt = sMessageBuilt & "Number: " & objItem.Numbers & ";"
        sMessageBuilt = sMessageBuilt & "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfos & ";"
        sMessageBuilt = sMessageBuilt & "PartNumber: " & objItem.PartNumbers & ";"
        sMessageBuilt = sMessageBuilt & "PMESignal: " & objItem.PMESignals & ";"
        sMessageBuilt = sMessageBuilt & "PoweredOn: " & objItem.PoweredOns & ";"
        sMessageBuilt = sMessageBuilt & "PurposeDescription: " & objItem.PurposeDescriptions & ";"
        sMessageBuilt = sMessageBuilt & "SerialNumber: " & objItem.SerialNumbers & ";"
        sMessageBuilt = sMessageBuilt & "Shared: " & objItem.Shareds & ";"
        sMessageBuilt = sMessageBuilt & "SKU: " & objItem.SKUs & ";"
        sMessageBuilt = sMessageBuilt & "SlotDesignation: " & objItem.SlotDesignations & ";"
        sMessageBuilt = sMessageBuilt & "SpecialPurpose: " & objItem.SpecialPurposes & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Statuss & ";"
        sMessageBuilt = sMessageBuilt & "SupportsHotPlug: " & objItem.SupportsHotPlugs & ";"
        sMessageBuilt = sMessageBuilt & "Tag: " & objItem.Tags & ";"
        sMessageBuilt = sMessageBuilt & "ThermalRating: " & objItem.ThermalRatings & ";"
        sMessageBuilt = sMessageBuilt & "VccMixedVoltageSupport: " & objItem.VccMixedVoltageSupports & ";"
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Versions & ";"
        sMessageBuilt = sMessageBuilt & "VppMixedVoltageSupport: " & objItem.VppMixedVoltageSupports
        
    
End Sub

Sub Win32_TemperatureProbe(objItem)
    'a little exhausting
    

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_TemperatureProbe", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Accuracy: " & objItem.Accuracy & ";"
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "CurrentReading: " & objItem.CurrentReading & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "IsLinear: " & objItem.IsLinear & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "LowerThresholdCritical: " & objItem.LowerThresholdCritical & ";"
        sMessageBuilt = sMessageBuilt & "LowerThresholdFatal: " & objItem.LowerThresholdFatal & ";"
        sMessageBuilt = sMessageBuilt & "LowerThresholdNonCritical: " & objItem.LowerThresholdNonCritical & ";"
        sMessageBuilt = sMessageBuilt & "MaxReadable: " & objItem.MaxReadable & ";"
        sMessageBuilt = sMessageBuilt & "MinReadable: " & objItem.MinReadable & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NominalReading: " & objItem.NominalReading & ";"
        sMessageBuilt = sMessageBuilt & "NormalMax: " & objItem.NormalMax & ";"
        sMessageBuilt = sMessageBuilt & "NormalMin: " & objItem.NormalMin & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "Resolution: " & objItem.Resolution & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        sMessageBuilt = sMessageBuilt & "Tolerance: " & objItem.Tolerance & ";"
        sMessageBuilt = sMessageBuilt & "UpperThresholdCritical: " & objItem.UpperThresholdCritical & ";"
        sMessageBuilt = sMessageBuilt & "UpperThresholdFatal: " & objItem.UpperThresholdFatal & ";"
        sMessageBuilt = sMessageBuilt & "UpperThresholdNonCritical: " & objItem.UpperThresholdNonCritical
        
    
End Sub

Sub Win32_USBController(objItem)
    

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_USBController", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MaxNumberControlled: " & objItem.MaxNumberControlled & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "ProtocolSupported: " & objItem.ProtocolSupported & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset
        
    
End Sub
Sub Win32_VideoConfiguration(objItem)
    

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoConfiguration", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "ActualColorResolution: " & objItem.ActualColorResolution & ";"
        sMessageBuilt = sMessageBuilt & "AdapterChipType: " & objItem.AdapterChipType & ";"
        sMessageBuilt = sMessageBuilt & "AdapterCompatibility: " & objItem.AdapterCompatibility & ";"
        sMessageBuilt = sMessageBuilt & "AdapterDACType: " & objItem.AdapterDACType & ";"
        sMessageBuilt = sMessageBuilt & "AdapterDescription: " & objItem.AdapterDescription & ";"
        sMessageBuilt = sMessageBuilt & "AdapterRAM: " & objItem.AdapterRAM & ";"
        sMessageBuilt = sMessageBuilt & "AdapterType: " & objItem.AdapterType & ";"
        sMessageBuilt = sMessageBuilt & "BitsPerPixel: " & objItem.BitsPerPixel & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ColorPlanes: " & objItem.ColorPlanes & ";"
        sMessageBuilt = sMessageBuilt & "ColorTableEntries: " & objItem.ColorTableEntries & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceSpecificPens: " & objItem.DeviceSpecificPens & ";"
        sMessageBuilt = sMessageBuilt & "DriverDate: " & objItem.DriverDate & ";"
        sMessageBuilt = sMessageBuilt & "HorizontalResolution: " & objItem.HorizontalResolution & ";"
        sMessageBuilt = sMessageBuilt & "InfFilename: " & objItem.InfFileName & ";"
        sMessageBuilt = sMessageBuilt & "InfSection: " & objItem.InfSection & ";"
        sMessageBuilt = sMessageBuilt & "InstalledDisplayDrivers: " & objItem.InstalledDisplayDrivers & ";"
        sMessageBuilt = sMessageBuilt & "MonitorManufacturer: " & objItem.MonitorManufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MonitorType: " & objItem.MonitorType & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "PixelsPerXLogicalInch: " & objItem.PixelsPerXLogicalInch & ";"
        sMessageBuilt = sMessageBuilt & "PixelsPerYLogicalInch: " & objItem.PixelsPerYLogicalInch & ";"
        sMessageBuilt = sMessageBuilt & "RefreshRate: " & objItem.RefreshRate & ";"
        sMessageBuilt = sMessageBuilt & "ScanMode: " & objItem.ScanMode & ";"
        sMessageBuilt = sMessageBuilt & "ScreenHeight: " & objItem.ScreenHeight & ";"
        sMessageBuilt = sMessageBuilt & "ScreenWidth: " & objItem.ScreenWidth & ";"
        sMessageBuilt = sMessageBuilt & "SettingID: " & objItem.SettingID & ";"
        sMessageBuilt = sMessageBuilt & "SystemPaletteEntries: " & objItem.SystemPaletteEntries & ";"
        sMessageBuilt = sMessageBuilt & "VerticalResolution: " & objItem.VerticalResolution
        
    
End Sub

Sub Win32_VideoController(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "AcceleratorCapabilities: " & objItem.AcceleratorCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "AdapterCompatibility: " & objItem.AdapterCompatibility & ";"
        sMessageBuilt = sMessageBuilt & "AdapterDACType: " & objItem.AdapterDACType & ";"
        sMessageBuilt = sMessageBuilt & "AdapterRAM: " & objItem.AdapterRAM & ";"
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "CapabilityDescriptions: " & objItem.CapabilityDescriptions & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ColorTableEntries: " & objItem.ColorTableEntries & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "CurrentBitsPerPixel: " & objItem.CurrentBitsPerPixel & ";"
        sMessageBuilt = sMessageBuilt & "CurrentHorizontalResolution: " & objItem.CurrentHorizontalResolution & ";"
        sMessageBuilt = sMessageBuilt & "CurrentNumberOfColors: " & objItem.CurrentNumberOfColors & ";"
        sMessageBuilt = sMessageBuilt & "CurrentNumberOfColumns: " & objItem.CurrentNumberOfColumns & ";"
        sMessageBuilt = sMessageBuilt & "CurrentNumberOfRows: " & objItem.CurrentNumberOfRows & ";"
        sMessageBuilt = sMessageBuilt & "CurrentRefreshRate: " & objItem.CurrentRefreshRate & ";"
        sMessageBuilt = sMessageBuilt & "CurrentScanMode: " & objItem.CurrentScanMode & ";"
        sMessageBuilt = sMessageBuilt & "CurrentVerticalResolution: " & objItem.CurrentVerticalResolution & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "DeviceSpecificPens: " & objItem.DeviceSpecificPens & ";"
        sMessageBuilt = sMessageBuilt & "DitherType: " & objItem.DitherType & ";"
        sMessageBuilt = sMessageBuilt & "DriverDate: " & objItem.DriverDate & ";"
        sMessageBuilt = sMessageBuilt & "DriverVersion: " & objItem.DriverVersion & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "ICMIntent: " & objItem.ICMIntent & ";"
        sMessageBuilt = sMessageBuilt & "ICMMethod: " & objItem.ICMMethod & ";"
        sMessageBuilt = sMessageBuilt & "InfFilename: " & objItem.InfFileName & ";"
        sMessageBuilt = sMessageBuilt & "InfSection: " & objItem.InfSection & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "InstalledDisplayDrivers: " & objItem.InstalledDisplayDrivers & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "MaxMemorySupported: " & objItem.MaxMemorySupported & ";"
        sMessageBuilt = sMessageBuilt & "MaxNumberControlled: " & objItem.MaxNumberControlled & ";"
        sMessageBuilt = sMessageBuilt & "MaxRefreshRate: " & objItem.MaxRefreshRate & ";"
        sMessageBuilt = sMessageBuilt & "MinRefreshRate: " & objItem.MinRefreshRate & ";"
        sMessageBuilt = sMessageBuilt & "Monochrome: " & objItem.Monochrome & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NumberOfColorPlanes: " & objItem.NumberOfColorPlanes & ";"
        sMessageBuilt = sMessageBuilt & "NumberOfVideoPages: " & objItem.NumberOfVideoPages & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "ProtocolSupported: " & objItem.ProtocolSupported & ";"
        sMessageBuilt = sMessageBuilt & "ReservedSystemPaletteEntries: " & objItem.ReservedSystemPaletteEntries & ";"
        sMessageBuilt = sMessageBuilt & "SpecificationVersion: " & objItem.SpecificationVersion & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        sMessageBuilt = sMessageBuilt & "SystemPaletteEntries: " & objItem.SystemPaletteEntries & ";"
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset & ";"
        sMessageBuilt = sMessageBuilt & "VideoArchitecture: " & objItem.VideoArchitecture & ";"
        sMessageBuilt = sMessageBuilt & "VideoMemoryType: " & objItem.VideoMemoryType & ";"
        sMessageBuilt = sMessageBuilt & "VideoMode: " & objItem.VideoMode & ";"
        sMessageBuilt = sMessageBuilt & "VideoModeDescription: " & objItem.VideoModeDescription & ";"
        sMessageBuilt = sMessageBuilt & "VideoProcessor: " & objItem.VideoProcessor
        

End Sub


Sub Win32_VoltageProbe(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_VoltageProbe", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Accuracy: " & objItem.Accuracy & ";"
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "CurrentReading: " & objItem.CurrentReading & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "IsLinear: " & objItem.IsLinear & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "LowerThresholdCritical: " & objItem.LowerThresholdCritical & ";"
        sMessageBuilt = sMessageBuilt & "LowerThresholdFatal: " & objItem.LowerThresholdFatal & ";"
        sMessageBuilt = sMessageBuilt & "LowerThresholdNonCritical: " & objItem.LowerThresholdNonCritical & ";"
        sMessageBuilt = sMessageBuilt & "MaxReadable: " & objItem.MaxReadable & ";"
        sMessageBuilt = sMessageBuilt & "MinReadable: " & objItem.MinReadable & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NominalReading: " & objItem.NominalReading & ";"
        sMessageBuilt = sMessageBuilt & "NormalMax: " & objItem.NormalMax & ";"
        sMessageBuilt = sMessageBuilt & "NormalMin: " & objItem.NormalMin & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "Resolution: " & objItem.Resolution & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        sMessageBuilt = sMessageBuilt & "Tolerance: " & objItem.Tolerance & ";"
        sMessageBuilt = sMessageBuilt & "UpperThresholdCritical: " & objItem.UpperThresholdCritical & ";"
        sMessageBuilt = sMessageBuilt & "UpperThresholdFatal: " & objItem.UpperThresholdFatal & ";"
        sMessageBuilt = sMessageBuilt & "UpperThresholdNonCritical: " & objItem.UpperThresholdNonCritical
        
    
End Sub

Sub Win32_1394Controller(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_1394Controller", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MaxNumberControlled: " & objItem.MaxNumberControlled & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "ProtocolSupported: " & objItem.ProtocolSupported & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset
        
    
End Sub

Sub Win32_BaseBoard(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigOptions: " & objItem.ConfigOptions
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Depth: " & objItem.Depth
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "Height: " & objItem.Height
        sMessageBuilt = sMessageBuilt & "HostingBoard: " & objItem.HostingBoard
        sMessageBuilt = sMessageBuilt & "HotSwappable: " & objItem.HotSwappable
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo
        sMessageBuilt = sMessageBuilt & "PartNumber: " & objItem.PartNumber
        sMessageBuilt = sMessageBuilt & "PoweredOn: " & objItem.PoweredOn
        sMessageBuilt = sMessageBuilt & "Product: " & objItem.Product
        sMessageBuilt = sMessageBuilt & "Removable: " & objItem.Removable
        sMessageBuilt = sMessageBuilt & "Replaceable: " & objItem.Replaceable
        sMessageBuilt = sMessageBuilt & "RequirementsDescription: " & objItem.RequirementsDescription
        sMessageBuilt = sMessageBuilt & "RequiresDaughterBoard: " & objItem.RequiresDaughterBoard
        sMessageBuilt = sMessageBuilt & "SerialNumber: " & objItem.SerialNumber
        sMessageBuilt = sMessageBuilt & "SKU: " & objItem.SKU
        sMessageBuilt = sMessageBuilt & "SlotLayout: " & objItem.SlotLayout
        sMessageBuilt = sMessageBuilt & "SpecialRequirements: " & objItem.SpecialRequirements
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "Tag: " & objItem.Tag
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Version
        sMessageBuilt = sMessageBuilt & "Weight: " & objItem.Weight
        sMessageBuilt = sMessageBuilt & "Width: " & objItem.Width
        
    
End Sub

Sub Win32_Battery(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_Battery", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "BatteryRechargeTime: " & objItem.BatteryRechargeTime
        sMessageBuilt = sMessageBuilt & "BatteryStatus: " & objItem.BatteryStatus
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "Chemistry: " & objItem.Chemistry
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DesignCapacity: " & objItem.DesignCapacity
        sMessageBuilt = sMessageBuilt & "DesignVoltage: " & objItem.DesignVoltage
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "EstimatedChargeRemaining: " & objItem.EstimatedChargeRemaining
        sMessageBuilt = sMessageBuilt & "EstimatedRunTime: " & objItem.EstimatedRunTime
        sMessageBuilt = sMessageBuilt & "ExpectedBatteryLife: " & objItem.ExpectedBatteryLife
        sMessageBuilt = sMessageBuilt & "ExpectedLife: " & objItem.ExpectedLife
        sMessageBuilt = sMessageBuilt & "FullChargeCapacity: " & objItem.FullChargeCapacity
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "MaxRechargeTime: " & objItem.MaxRechargeTime
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "SmartBatteryVersion: " & objItem.SmartBatteryVersion
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "TimeOnBattery: " & objItem.TimeOnBattery
        sMessageBuilt = sMessageBuilt & "TimeToFullCharge: " & objItem.TimeToFullCharge
        
    
End Sub

Sub Win32_BIOS(objItem)

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "BiosCharacteristics: " & objItem.BiosCharacteristics
        sMessageBuilt = sMessageBuilt & "BuildNumber: " & objItem.BuildNumber
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "CodeSet: " & objItem.CodeSet
        sMessageBuilt = sMessageBuilt & "CurrentLanguage: " & objItem.CurrentLanguage
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "IdentificationCode: " & objItem.IdentificationCode
        sMessageBuilt = sMessageBuilt & "InstallableLanguages: " & objItem.InstallableLanguages
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LanguageEdition: " & objItem.LanguageEdition
        sMessageBuilt = sMessageBuilt & "ListOfLanguages: " & objItem.ListOfLanguages
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "OtherTargetOS: " & objItem.OtherTargetOS
        sMessageBuilt = sMessageBuilt & "PrimaryBIOS: " & objItem.PrimaryBIOS
        sMessageBuilt = sMessageBuilt & "ReleaseDate: " & objItem.ReleaseDate
        sMessageBuilt = sMessageBuilt & "SerialNumber: " & objItem.SerialNumber
        sMessageBuilt = sMessageBuilt & "SMBIOSBIOSVersion: " & objItem.SMBIOSBIOSVersion
        sMessageBuilt = sMessageBuilt & "SMBIOSMajorVersion: " & objItem.SMBIOSMajorVersion
        sMessageBuilt = sMessageBuilt & "SMBIOSMinorVersion: " & objItem.SMBIOSMinorVersion
        sMessageBuilt = sMessageBuilt & "SMBIOSPresent: " & objItem.SMBIOSPresent
        sMessageBuilt = sMessageBuilt & "SoftwareElementID: " & objItem.SoftwareElementID
        sMessageBuilt = sMessageBuilt & "SoftwareElementState: " & objItem.SoftwareElementState
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "TargetOperatingSystem: " & objItem.TargetOperatingSystem
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Version
        
    
End Sub

Sub Win32_Bus(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_Bus", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "BusNum: " & objItem.BusNum
        sMessageBuilt = sMessageBuilt & "BusType: " & objItem.BusType
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        
    
End Sub

Sub Win32_CacheMemory(objItem)

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_CacheMemory", , 48)
        
        
        sMessageBuilt = sMessageBuilt & "Access: " & objItem.Access
        sMessageBuilt = sMessageBuilt & "AdditionalErrorData: " & objItem.AdditionalErrorData
        sMessageBuilt = sMessageBuilt & "Associativity: " & objItem.Associativity
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "BlockSize: " & objItem.BlockSize
        sMessageBuilt = sMessageBuilt & "CacheSpeed: " & objItem.CacheSpeed
        sMessageBuilt = sMessageBuilt & "CacheType: " & objItem.CacheType
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CorrectableError: " & objItem.CorrectableError
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        'sMessageBuilt = sMessageBuilt & "CurrentSRAM: " & objItem.CurrentSRAM
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "EndingAddress: " & objItem.EndingAddress
        sMessageBuilt = sMessageBuilt & "ErrorAccess: " & objItem.ErrorAccess
        sMessageBuilt = sMessageBuilt & "ErrorAddress: " & objItem.ErrorAddress
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorCorrectType: " & objItem.ErrorCorrectType
        sMessageBuilt = sMessageBuilt & "ErrorData: " & objItem.ErrorData
        sMessageBuilt = sMessageBuilt & "ErrorDataOrder: " & objItem.ErrorDataOrder
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "ErrorInfo: " & objItem.ErrorInfo
        sMessageBuilt = sMessageBuilt & "ErrorMethodology: " & objItem.ErrorMethodology
        sMessageBuilt = sMessageBuilt & "ErrorResolution: " & objItem.ErrorResolution
        sMessageBuilt = sMessageBuilt & "ErrorTime: " & objItem.ErrorTime
        sMessageBuilt = sMessageBuilt & "ErrorTransferSize: " & objItem.ErrorTransferSize
        sMessageBuilt = sMessageBuilt & "FlushTimer: " & objItem.FlushTimer
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "InstalledSize: " & objItem.InstalledSize
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Level: " & objItem.Level
        sMessageBuilt = sMessageBuilt & "LineSize: " & objItem.LineSize
        sMessageBuilt = sMessageBuilt & "Location: " & objItem.Location
        sMessageBuilt = sMessageBuilt & "MaxCacheSize: " & objItem.MaxCacheSize
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "NumberOfBlocks: " & objItem.NumberOfBlocks
        sMessageBuilt = sMessageBuilt & "OtherErrorDescription: " & objItem.OtherErrorDescription
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "Purpose: " & objItem.Purpose
        sMessageBuilt = sMessageBuilt & "ReadPolicy: " & objItem.ReadPolicy
        sMessageBuilt = sMessageBuilt & "ReplacementPolicy: " & objItem.ReplacementPolicy
        sMessageBuilt = sMessageBuilt & "StartingAddress: " & objItem.StartingAddress
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        'sMessageBuilt = sMessageBuilt & "SupportedSRAM: " & objItem.SupportedSRAM
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemLevelAddress: " & objItem.SystemLevelAddress
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "WritePolicy: " & objItem.WritePolicy
        
    
End Sub

Sub Win32_CDROMDrive(objItem)

On Error Resume Next
    
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_CDROMDrive", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Capabilities: " & objItem.Capabilities
        sMessageBuilt = sMessageBuilt & "CapabilityDescriptions: " & objItem.CapabilityDescriptions
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "CompressionMethod: " & objItem.CompressionMethod
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "DefaultBlockSize: " & objItem.DefaultBlockSize
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "Drive: " & objItem.Drive
        sMessageBuilt = sMessageBuilt & "DriveIntegrity: " & objItem.DriveIntegrity
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "ErrorMethodology: " & objItem.ErrorMethodology
        sMessageBuilt = sMessageBuilt & "FileSystemFlags: " & objItem.FileSystemFlags
        sMessageBuilt = sMessageBuilt & "FileSystemFlagsEx: " & objItem.FileSystemFlagsEx
        sMessageBuilt = sMessageBuilt & "Id: " & objItem.Id
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "MaxBlockSize: " & objItem.MaxBlockSize
        sMessageBuilt = sMessageBuilt & "MaximumComponentLength: " & objItem.MaximumComponentLength
        sMessageBuilt = sMessageBuilt & "MaxMediaSize: " & objItem.MaxMediaSize
        sMessageBuilt = sMessageBuilt & "MediaLoaded: " & objItem.MediaLoaded
        sMessageBuilt = sMessageBuilt & "MediaType: " & objItem.MediaType
        sMessageBuilt = sMessageBuilt & "MinBlockSize: " & objItem.MinBlockSize
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "NeedsCleaning: " & objItem.NeedsCleaning
        sMessageBuilt = sMessageBuilt & "NumberOfMediaSupported: " & objItem.NumberOfMediaSupported
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "RevisionLevel: " & objItem.RevisionLevel
        sMessageBuilt = sMessageBuilt & "SCSIBus: " & objItem.SCSIBus
        sMessageBuilt = sMessageBuilt & "SCSILogicalUnit: " & objItem.SCSILogicalUnit
        sMessageBuilt = sMessageBuilt & "SCSIPort: " & objItem.SCSIPort
        sMessageBuilt = sMessageBuilt & "SCSITargetId: " & objItem.SCSITargetId
        sMessageBuilt = sMessageBuilt & "Size: " & objItem.Size
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "TransferRate: " & objItem.TransferRate
        sMessageBuilt = sMessageBuilt & "VolumeName: " & objItem.VolumeName
        sMessageBuilt = sMessageBuilt & "VolumeSerialNumber: " & objItem.VolumeSerialNumber
        
    
End Sub


Sub Win32_CurrentProbe(objItem)

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_CurrentProbe", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Accuracy: " & objItem.Accuracy
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "CurrentReading: " & objItem.CurrentReading
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "IsLinear: " & objItem.IsLinear
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "LowerThresholdCritical: " & objItem.LowerThresholdCritical
        sMessageBuilt = sMessageBuilt & "LowerThresholdFatal: " & objItem.LowerThresholdFatal
        sMessageBuilt = sMessageBuilt & "LowerThresholdNonCritical: " & objItem.LowerThresholdNonCritical
        sMessageBuilt = sMessageBuilt & "MaxReadable: " & objItem.MaxReadable
        sMessageBuilt = sMessageBuilt & "MinReadable: " & objItem.MinReadable
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "NominalReading: " & objItem.NominalReading
        sMessageBuilt = sMessageBuilt & "NormalMax: " & objItem.NormalMax
        sMessageBuilt = sMessageBuilt & "NormalMin: " & objItem.NormalMin
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "Resolution: " & objItem.Resolution
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "Tolerance: " & objItem.Tolerance
        sMessageBuilt = sMessageBuilt & "UpperThresholdCritical: " & objItem.UpperThresholdCritical
        sMessageBuilt = sMessageBuilt & "UpperThresholdFatal: " & objItem.UpperThresholdFatal
        sMessageBuilt = sMessageBuilt & "UpperThresholdNonCritical: " & objItem.UpperThresholdNonCritical
        
    
End Sub


Sub Win32_DesktopMonitor(objItem)

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Bandwidth: " & objItem.Bandwidth
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "DisplayType: " & objItem.DisplayType
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "IsLocked: " & objItem.IsLocked
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "MonitorManufacturer: " & objItem.MonitorManufacturer
        sMessageBuilt = sMessageBuilt & "MonitorType: " & objItem.MonitorType
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PixelsPerXLogicalInch: " & objItem.PixelsPerXLogicalInch
        sMessageBuilt = sMessageBuilt & "PixelsPerYLogicalInch: " & objItem.PixelsPerYLogicalInch
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "ScreenHeight: " & objItem.ScreenHeight
        sMessageBuilt = sMessageBuilt & "ScreenWidth: " & objItem.ScreenWidth
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        
    

End Sub


Sub Win32_DeviceMemoryAddress(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_DeviceMemoryAddress", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "CSCreationClassName: " & objItem.CSCreationClassName
        sMessageBuilt = sMessageBuilt & "CSName: " & objItem.CSName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "EndingAddress: " & objItem.EndingAddress
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "MemoryType: " & objItem.MemoryType
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "StartingAddress: " & objItem.StartingAddress
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        
    

End Sub


Sub Win32_DMAChannel(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_DMAChannel", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "AddressSize: " & objItem.AddressSize
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "BurstMode: " & objItem.BurstMode
        sMessageBuilt = sMessageBuilt & "ByteMode: " & objItem.ByteMode
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ChannelTiming: " & objItem.ChannelTiming
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "CSCreationClassName: " & objItem.CSCreationClassName
        sMessageBuilt = sMessageBuilt & "CSName: " & objItem.CSName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DMAChannel: " & objItem.DMAChannel
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "MaxTransferSize: " & objItem.MaxTransferSize
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "Port: " & objItem.Port
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "TransferWidths: " & objItem.TransferWidths
        sMessageBuilt = sMessageBuilt & "TypeCTiming: " & objItem.TypeCTiming
        sMessageBuilt = sMessageBuilt & "WordMode: " & objItem.WordMode
        
    

End Sub


Sub Win32_Fan(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_Fan", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "ActiveCooling: " & objItem.ActiveCooling
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DesiredSpeed: " & objItem.DesiredSpeed
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "VariableSpeed: " & objItem.VariableSpeed
        
    

End Sub


Sub Win32_FloppyController(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_FloppyController", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "MaxNumberControlled: " & objItem.MaxNumberControlled
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "ProtocolSupported: " & objItem.ProtocolSupported
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset
        
    

End Sub


Sub Win32_FloppyDrive(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_FloppyDrive", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Capabilities: " & objItem.Capabilities
        sMessageBuilt = sMessageBuilt & "CapabilityDescriptions: " & objItem.CapabilityDescriptions
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "CompressionMethod: " & objItem.CompressionMethod
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "DefaultBlockSize: " & objItem.DefaultBlockSize
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "ErrorMethodology: " & objItem.ErrorMethodology
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "MaxBlockSize: " & objItem.MaxBlockSize
        sMessageBuilt = sMessageBuilt & "MaxMediaSize: " & objItem.MaxMediaSize
        sMessageBuilt = sMessageBuilt & "MinBlockSize: " & objItem.MinBlockSize
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "NeedsCleaning: " & objItem.NeedsCleaning
        sMessageBuilt = sMessageBuilt & "NumberOfMediaSupported: " & objItem.NumberOfMediaSupported
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        
    

End Sub


Sub Win32_IDEController(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_IDEController", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "MaxNumberControlled: " & objItem.MaxNumberControlled
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "ProtocolSupported: " & objItem.ProtocolSupported
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset
        
    

End Sub


Sub Win32_IRQResource(objItem)

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_IRQResource", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "CSCreationClassName: " & objItem.CSCreationClassName
        sMessageBuilt = sMessageBuilt & "CSName: " & objItem.CSName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "Hardware: " & objItem.Hardware
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "IRQNumber: " & objItem.IRQNumber
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "Shareable: " & objItem.Shareable
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "TriggerLevel: " & objItem.TriggerLevel
        sMessageBuilt = sMessageBuilt & "TriggerType: " & objItem.TriggerType
        sMessageBuilt = sMessageBuilt & "Vector: " & objItem.Vector
        
    

End Sub


Sub Win32_Keyboard(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_Keyboard", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "IsLocked: " & objItem.IsLocked
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Layout: " & objItem.Layout
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "NumberOfFunctionKeys: " & objItem.NumberOfFunctionKeys
        sMessageBuilt = sMessageBuilt & "Password: " & objItem.Password
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        
    

End Sub


Sub Win32_MemoryArray(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_MemoryArray", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Access: " & objItem.Access & ";"
        sMessageBuilt = sMessageBuilt & "AdditionalErrorData: " & objItem.AdditionalErrorData & ";"
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "BlockSize: " & objItem.BlockSize & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CorrectableError: " & objItem.CorrectableError & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "EndingAddress: " & objItem.EndingAddress & ";"
        sMessageBuilt = sMessageBuilt & "ErrorAccess: " & objItem.ErrorAccess & ";"
        sMessageBuilt = sMessageBuilt & "ErrorAddress: " & objItem.ErrorAddress & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorData: " & objItem.ErrorData & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDataOrder: " & objItem.ErrorDataOrder & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "ErrorGranularity: " & objItem.ErrorGranularity & ";"
        sMessageBuilt = sMessageBuilt & "ErrorInfo: " & objItem.ErrorInfo & ";"
        sMessageBuilt = sMessageBuilt & "ErrorMethodology: " & objItem.ErrorMethodology & ";"
        sMessageBuilt = sMessageBuilt & "ErrorResolution: " & objItem.ErrorResolution & ";"
        sMessageBuilt = sMessageBuilt & "ErrorTime: " & objItem.ErrorTime & ";"
        sMessageBuilt = sMessageBuilt & "ErrorTransferSize: " & objItem.ErrorTransferSize & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NumberOfBlocks: " & objItem.NumberOfBlocks & ";"
        sMessageBuilt = sMessageBuilt & "OtherErrorDescription: " & objItem.OtherErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "Purpose: " & objItem.Purpose & ";"
        sMessageBuilt = sMessageBuilt & "StartingAddress: " & objItem.StartingAddress & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemLevelAddress: " & objItem.SystemLevelAddress & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        
    

End Sub


Sub Win32_MemoryDevice(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_MemoryDevice", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Access: " & objItem.Access & ";"
        sMessageBuilt = sMessageBuilt & "AdditionalErrorData: " & objItem.AdditionalErrorData & ";"
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "BlockSize: " & objItem.BlockSize & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CorrectableError: " & objItem.CorrectableError & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "EndingAddress: " & objItem.EndingAddress & ";"
        sMessageBuilt = sMessageBuilt & "ErrorAccess: " & objItem.ErrorAccess & ";"
        sMessageBuilt = sMessageBuilt & "ErrorAddress: " & objItem.ErrorAddress & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorData: " & objItem.ErrorData & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDataOrder: " & objItem.ErrorDataOrder & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "ErrorGranularity: " & objItem.ErrorGranularity & ";"
        sMessageBuilt = sMessageBuilt & "ErrorInfo: " & objItem.ErrorInfo & ";"
        sMessageBuilt = sMessageBuilt & "ErrorMethodology: " & objItem.ErrorMethodology & ";"
        sMessageBuilt = sMessageBuilt & "ErrorResolution: " & objItem.ErrorResolution & ";"
        sMessageBuilt = sMessageBuilt & "ErrorTime: " & objItem.ErrorTime & ";"
        sMessageBuilt = sMessageBuilt & "ErrorTransferSize: " & objItem.ErrorTransferSize & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NumberOfBlocks: " & objItem.NumberOfBlocks & ";"
        sMessageBuilt = sMessageBuilt & "OtherErrorDescription: " & objItem.OtherErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "Purpose: " & objItem.Purpose & ";"
        sMessageBuilt = sMessageBuilt & "StartingAddress: " & objItem.StartingAddress & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemLevelAddress: " & objItem.SystemLevelAddress & ";"
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName & ";"
        
    

End Sub


Sub Win32_MotherboardDevice(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_MotherboardDevice", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "PrimaryBusType: " & objItem.PrimaryBusType
        sMessageBuilt = sMessageBuilt & "RevisionNumber: " & objItem.RevisionNumber
        sMessageBuilt = sMessageBuilt & "SecondaryBusType: " & objItem.SecondaryBusType
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        
    

End Sub


Sub Win32_NetworkAdapter(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "AdapterType$$ " & objItem.AdapterType & ";"
        sMessageBuilt = sMessageBuilt & "AutoSense$$ " & objItem.AutoSense & ";"
        sMessageBuilt = sMessageBuilt & "Availability$$ " & objItem.Availability & ";"
        sMessageBuilt = sMessageBuilt & "Caption$$ " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode$$ " & objItem.ConfigManagerErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig$$ " & objItem.ConfigManagerUserConfig & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName$$ " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "Description$$ " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceID$$ " & objItem.DeviceID & ";"
        sMessageBuilt = sMessageBuilt & "ErrorCleared$$ " & objItem.ErrorCleared & ";"
        sMessageBuilt = sMessageBuilt & "ErrorDescription$$ " & objItem.ErrorDescription & ";"
        sMessageBuilt = sMessageBuilt & "Index$$ " & objItem.Index & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate$$ " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "Installed$$ " & objItem.Installed & ";"
        sMessageBuilt = sMessageBuilt & "LastErrorCode$$ " & objItem.LastErrorCode & ";"
        sMessageBuilt = sMessageBuilt & "MACAddress$$ " & objItem.MACAddress & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer$$ " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MaxNumberControlled$$ " & objItem.MaxNumberControlled & ";"
        sMessageBuilt = sMessageBuilt & "MaxSpeed$$ " & objItem.MaxSpeed & ";"
        sMessageBuilt = sMessageBuilt & "Name$$ " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "NetworkAddresses$$ " & objItem.NetworkAddresses & ";"
        sMessageBuilt = sMessageBuilt & "PermanentAddress$$ " & objItem.PermanentAddress & ";"
        sMessageBuilt = sMessageBuilt & "PNPDeviceID$$ " & objItem.PNPDeviceID & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities$$ " & objItem.PowerManagementCapabilities & ";"
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported$$ " & objItem.PowerManagementSupported & ";"
        sMessageBuilt = sMessageBuilt & "ProductName$$ " & objItem.ProductName & ";"
        sMessageBuilt = sMessageBuilt & "ServiceName$$ " & objItem.ServiceName & ";"
        sMessageBuilt = sMessageBuilt & "Speed$$ " & objItem.Speed & ";"
        sMessageBuilt = sMessageBuilt & "Status$$ " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "StatusInfo$$ " & objItem.StatusInfo & ";"
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName$$ " & objItem.SystemCreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "SystemName$$ " & objItem.SystemName & ";"
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset$$ " & objItem.TimeOfLastReset & ";$"
        
    

End Sub


Sub Win32_NetworkAdapterConfiguration(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "ArpAlwaysSourceRoute$$ " & objItem.ArpAlwaysSourceRoute & ";"
        sMessageBuilt = sMessageBuilt & "ArpUseEtherSNAP$$ " & objItem.ArpUseEtherSNAP & ";"
        sMessageBuilt = sMessageBuilt & "Caption$$ " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "DatabasePath$$ " & objItem.DatabasePath & ";"
        sMessageBuilt = sMessageBuilt & "DeadGWDetectEnabled$$ " & objItem.DeadGWDetectEnabled & ";"
        sMessageBuilt = sMessageBuilt & "DefaultIPGateway$$ " & objItem.DefaultIPGateway & ";"
        sMessageBuilt = sMessageBuilt & "DefaultTOS$$ " & objItem.DefaultTOS & ";"
        sMessageBuilt = sMessageBuilt & "DefaultTTL$$ " & objItem.DefaultTTL & ";"
        sMessageBuilt = sMessageBuilt & "Description$$ " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DHCPEnabled$$ " & objItem.DHCPEnabled & ";"
        sMessageBuilt = sMessageBuilt & "DHCPLeaseExpires$$ " & objItem.DHCPLeaseExpires & ";"
        sMessageBuilt = sMessageBuilt & "DHCPLeaseObtained$$ " & objItem.DHCPLeaseObtained & ";"
        sMessageBuilt = sMessageBuilt & "DHCPServer$$ " & objItem.DHCPServer & ";"
        sMessageBuilt = sMessageBuilt & "DNSDomain$$ " & objItem.DNSDomain & ";"
        sMessageBuilt = sMessageBuilt & "DNSDomainSuffixSearchOrder$$ " & objItem.DNSDomainSuffixSearchOrder & ";"
        sMessageBuilt = sMessageBuilt & "DNSEnabledForWINSResolution$$ " & objItem.DNSEnabledForWINSResolution & ";"
        sMessageBuilt = sMessageBuilt & "DNSHostName$$ " & objItem.DNSHostName & ";"
        sMessageBuilt = sMessageBuilt & "DNSServerSearchOrder$$ " & objItem.DNSServerSearchOrder & ";"
        sMessageBuilt = sMessageBuilt & "DomainDNSRegistrationEnabled$$ " & objItem.DomainDNSRegistrationEnabled & ";"
        sMessageBuilt = sMessageBuilt & "ForwardBufferMemory$$ " & objItem.ForwardBufferMemory & ";"
        sMessageBuilt = sMessageBuilt & "FullDNSRegistrationEnabled$$ " & objItem.FullDNSRegistrationEnabled & ";"
        sMessageBuilt = sMessageBuilt & "GatewayCostMetric$$ " & objItem.GatewayCostMetric & ";"
        sMessageBuilt = sMessageBuilt & "IGMPLevel$$ " & objItem.IGMPLevel & ";"
        sMessageBuilt = sMessageBuilt & "Index$$ " & objItem.Index & ";"
        sMessageBuilt = sMessageBuilt & "IPAddress$$ " & objItem.IPAddress & ";"
        sMessageBuilt = sMessageBuilt & "IPConnectionMetric$$ " & objItem.IPConnectionMetric & ";"
        sMessageBuilt = sMessageBuilt & "IPEnabled$$ " & objItem.IPEnabled & ";"
        sMessageBuilt = sMessageBuilt & "IPFilterSecurityEnabled$$ " & objItem.IPFilterSecurityEnabled & ";"
        sMessageBuilt = sMessageBuilt & "IPPortSecurityEnabled$$ " & objItem.IPPortSecurityEnabled & ";"
        sMessageBuilt = sMessageBuilt & "IPSecPermitIPProtocols$$ " & objItem.IPSecPermitIPProtocols & ";"
        sMessageBuilt = sMessageBuilt & "IPSecPermitTCPPorts$$ " & objItem.IPSecPermitTCPPorts & ";"
        sMessageBuilt = sMessageBuilt & "IPSecPermitUDPPorts$$ " & objItem.IPSecPermitUDPPorts & ";"
        sMessageBuilt = sMessageBuilt & "IPSubnet$$ " & objItem.IPSubnet & ";"
        sMessageBuilt = sMessageBuilt & "IPUseZeroBroadcast$$ " & objItem.IPUseZeroBroadcast & ";"
        sMessageBuilt = sMessageBuilt & "IPXAddress$$ " & objItem.IPXAddress & ";"
        sMessageBuilt = sMessageBuilt & "IPXEnabled$$ " & objItem.IPXEnabled & ";"
        sMessageBuilt = sMessageBuilt & "IPXFrameType$$ " & objItem.IPXFrameType & ";"
        sMessageBuilt = sMessageBuilt & "IPXMediaType$$ " & objItem.IPXMediaType & ";"
        sMessageBuilt = sMessageBuilt & "IPXNetworkNumber$$ " & objItem.IPXNetworkNumber & ";"
        sMessageBuilt = sMessageBuilt & "IPXVirtualNetNumber$$ " & objItem.IPXVirtualNetNumber & ";"
        sMessageBuilt = sMessageBuilt & "KeepAliveInterval$$ " & objItem.KeepAliveInterval & ";"
        sMessageBuilt = sMessageBuilt & "KeepAliveTime$$ " & objItem.KeepAliveTime & ";"
        sMessageBuilt = sMessageBuilt & "MACAddress$$ " & objItem.MACAddress & ";"
        sMessageBuilt = sMessageBuilt & "MTU$$ " & objItem.MTU & ";"
        sMessageBuilt = sMessageBuilt & "NumForwardPackets$$ " & objItem.NumForwardPackets & ";"
        sMessageBuilt = sMessageBuilt & "PMTUBHDetectEnabled$$ " & objItem.PMTUBHDetectEnabled & ";"
        sMessageBuilt = sMessageBuilt & "PMTUDiscoveryEnabled$$ " & objItem.PMTUDiscoveryEnabled & ";"
        sMessageBuilt = sMessageBuilt & "ServiceName$$ " & objItem.ServiceName & ";"
        sMessageBuilt = sMessageBuilt & "SettingID$$ " & objItem.SettingID & ";"
        sMessageBuilt = sMessageBuilt & "TcpipNetbiosOptions$$ " & objItem.TcpipNetbiosOptions & ";"
        sMessageBuilt = sMessageBuilt & "TcpMaxConnectRetransmissions$$ " & objItem.TcpMaxConnectRetransmissions & ";"
        sMessageBuilt = sMessageBuilt & "TcpMaxDataRetransmissions$$ " & objItem.TcpMaxDataRetransmissions & ";"
        sMessageBuilt = sMessageBuilt & "TcpNumConnections$$ " & objItem.TcpNumConnections & ";"
        sMessageBuilt = sMessageBuilt & "TcpUseRFC1122UrgentPointer$$ " & objItem.TcpUseRFC1122UrgentPointer & ";"
        sMessageBuilt = sMessageBuilt & "TcpWindowSize$$ " & objItem.TcpWindowSize & ";"
        sMessageBuilt = sMessageBuilt & "WINSEnableLMHostsLookup$$ " & objItem.WINSEnableLMHostsLookup & ";"
        sMessageBuilt = sMessageBuilt & "WINSHostLookupFile$$ " & objItem.WINSHostLookupFile & ";"
        sMessageBuilt = sMessageBuilt & "WINSPrimaryServer$$ " & objItem.WINSPrimaryServer & ";"
        sMessageBuilt = sMessageBuilt & "WINSScopeID$$ " & objItem.WINSScopeID & ";"
        sMessageBuilt = sMessageBuilt & "WINSSecondaryServer$$ " & objItem.WINSSecondaryServer & ";$"
        
    

End Sub


Sub Win32_OnBoardDevice(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_OnBoardDevice", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceType: " & objItem.DeviceType
        sMessageBuilt = sMessageBuilt & "Enabled: " & objItem.Enabled
        sMessageBuilt = sMessageBuilt & "HotSwappable: " & objItem.HotSwappable
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo
        sMessageBuilt = sMessageBuilt & "PartNumber: " & objItem.PartNumber
        sMessageBuilt = sMessageBuilt & "PoweredOn: " & objItem.PoweredOn
        sMessageBuilt = sMessageBuilt & "Removable: " & objItem.Removable
        sMessageBuilt = sMessageBuilt & "Replaceable: " & objItem.Replaceable
        sMessageBuilt = sMessageBuilt & "SerialNumber: " & objItem.SerialNumber
        sMessageBuilt = sMessageBuilt & "SKU: " & objItem.SKU
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "Tag: " & objItem.Tag
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Version
        
    

End Sub


Sub Win32_ParallelPort(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_ParallelPort", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Capabilities: " & objItem.Capabilities
        sMessageBuilt = sMessageBuilt & "CapabilityDescriptions: " & objItem.CapabilityDescriptions
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "DMASupport: " & objItem.DMASupport
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "MaxNumberControlled: " & objItem.MaxNumberControlled
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "OSAutoDiscovered: " & objItem.OSAutoDiscovered
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "ProtocolSupported: " & objItem.ProtocolSupported
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset
        
    

End Sub


Sub Win32_PhysicalMemory(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "BankLabel: " & objItem.BankLabel & ";"
        sMessageBuilt = sMessageBuilt & "Capacity: " & objItem.Capacity & ";"
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "DataWidth: " & objItem.DataWidth & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "DeviceLocator: " & objItem.DeviceLocator & ";"
        sMessageBuilt = sMessageBuilt & "FormFactor: " & objItem.FormFactor & ";"
        sMessageBuilt = sMessageBuilt & "HotSwappable: " & objItem.HotSwappable & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "InterleaveDataDepth: " & objItem.InterleaveDataDepth & ";"
        sMessageBuilt = sMessageBuilt & "InterleavePosition: " & objItem.InterleavePosition & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MemoryType: " & objItem.MemoryType & ";"
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo & ";"
        sMessageBuilt = sMessageBuilt & "PartNumber: " & objItem.PartNumber & ";"
        sMessageBuilt = sMessageBuilt & "PositionInRow: " & objItem.PositionInRow & ";"
        sMessageBuilt = sMessageBuilt & "PoweredOn: " & objItem.PoweredOn & ";"
        sMessageBuilt = sMessageBuilt & "Removable: " & objItem.Removable & ";"
        sMessageBuilt = sMessageBuilt & "Replaceable: " & objItem.Replaceable & ";"
        sMessageBuilt = sMessageBuilt & "SerialNumber: " & objItem.SerialNumber & ";"
        sMessageBuilt = sMessageBuilt & "SKU: " & objItem.SKU & ";"
        sMessageBuilt = sMessageBuilt & "Speed: " & objItem.Speed & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "Tag: " & objItem.Tag & ";"
        sMessageBuilt = sMessageBuilt & "TotalWidth: " & objItem.TotalWidth & ";"
        sMessageBuilt = sMessageBuilt & "TypeDetail: " & objItem.TypeDetail & ";"
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Version & ";"
        
    

End Sub


Sub Win32_PhysicalMemoryArray(objItem)

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption & ";"
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName & ";"
        sMessageBuilt = sMessageBuilt & "Depth: " & objItem.Depth & ";"
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description & ";"
        sMessageBuilt = sMessageBuilt & "Height: " & objItem.Height & ";"
        sMessageBuilt = sMessageBuilt & "HotSwappable: " & objItem.HotSwappable & ";"
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate & ";"
        sMessageBuilt = sMessageBuilt & "Location: " & objItem.Location & ";"
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer & ";"
        sMessageBuilt = sMessageBuilt & "MaxCapacity: " & objItem.MaxCapacity & ";"
        sMessageBuilt = sMessageBuilt & "MemoryDevices: " & objItem.MemoryDevices & ";"
        sMessageBuilt = sMessageBuilt & "MemoryErrorCorrection: " & objItem.MemoryErrorCorrection & ";"
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name & ";"
        sMessageBuilt = sMessageBuilt & "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo & ";"
        sMessageBuilt = sMessageBuilt & "PartNumber: " & objItem.PartNumber & ";"
        sMessageBuilt = sMessageBuilt & "PoweredOn: " & objItem.PoweredOn & ";"
        sMessageBuilt = sMessageBuilt & "Removable: " & objItem.Removable & ";"
        sMessageBuilt = sMessageBuilt & "Replaceable: " & objItem.Replaceable & ";"
        sMessageBuilt = sMessageBuilt & "SerialNumber: " & objItem.SerialNumber & ";"
        sMessageBuilt = sMessageBuilt & "SKU: " & objItem.SKU & ";"
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status & ";"
        sMessageBuilt = sMessageBuilt & "Tag: " & objItem.Tag & ";"
        sMessageBuilt = sMessageBuilt & "Use: " & objItem.Use & ";"
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Version & ";"
        sMessageBuilt = sMessageBuilt & "Weight: " & objItem.Weight & ";"
        sMessageBuilt = sMessageBuilt & "Width: " & objItem.Width & ";"
        
    

End Sub


Sub Win32_PnPEntity(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ClassGuid: " & objItem.ClassGuid
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "Service: " & objItem.Service
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        
    

End Sub


Sub Win32_PointingDevice(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PointingDevice", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "DeviceInterface: " & objItem.DeviceInterface
        sMessageBuilt = sMessageBuilt & "DoubleSpeedThreshold: " & objItem.DoubleSpeedThreshold
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "Handedness: " & objItem.Handedness
        sMessageBuilt = sMessageBuilt & "HardwareType: " & objItem.HardwareType
        sMessageBuilt = sMessageBuilt & "InfFileName: " & objItem.InfFileName
        sMessageBuilt = sMessageBuilt & "InfSection: " & objItem.InfSection
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "IsLocked: " & objItem.IsLocked
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "NumberOfButtons: " & objItem.NumberOfButtons
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PointingType: " & objItem.PointingType
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "QuadSpeedThreshold: " & objItem.QuadSpeedThreshold
        sMessageBuilt = sMessageBuilt & "Resolution: " & objItem.Resolution
        sMessageBuilt = sMessageBuilt & "SampleRate: " & objItem.SampleRate
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "Synch: " & objItem.Synch
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        
    

End Sub


Sub Win32_PortConnector(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PortConnector", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConnectorPinout: " & objItem.ConnectorPinout
        sMessageBuilt = sMessageBuilt & "ConnectorType: " & objItem.ConnectorType
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "ExternalReferenceDesignator: " & objItem.ExternalReferenceDesignator
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "InternalReferenceDesignator: " & objItem.InternalReferenceDesignator
        sMessageBuilt = sMessageBuilt & "Manufacturer: " & objItem.Manufacturer
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo
        sMessageBuilt = sMessageBuilt & "PartNumber: " & objItem.PartNumber
        sMessageBuilt = sMessageBuilt & "PortType: " & objItem.PortType
        sMessageBuilt = sMessageBuilt & "PoweredOn: " & objItem.PoweredOn
        sMessageBuilt = sMessageBuilt & "SerialNumber: " & objItem.SerialNumber
        sMessageBuilt = sMessageBuilt & "SKU: " & objItem.SKU
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "Tag: " & objItem.Tag
        sMessageBuilt = sMessageBuilt & "Version: " & objItem.Version
        
    

End Sub


Sub Win32_PortResource(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PortResource", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Alias: " & objItem.Alias
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "CSCreationClassName: " & objItem.CSCreationClassName
        sMessageBuilt = sMessageBuilt & "CSName: " & objItem.CSName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "EndingAddress: " & objItem.EndingAddress
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "StartingAddress: " & objItem.StartingAddress
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        
    

End Sub


Sub Win32_POTSModem(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_POTSModem", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "AnswerMode: " & objItem.AnswerMode
        sMessageBuilt = sMessageBuilt & "AttachedTo: " & objItem.AttachedTo
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "BlindOff: " & objItem.BlindOff
        sMessageBuilt = sMessageBuilt & "BlindOn: " & objItem.BlindOn
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "CompatibilityFlags: " & objItem.CompatibilityFlags
        sMessageBuilt = sMessageBuilt & "CompressionInfo: " & objItem.CompressionInfo
        sMessageBuilt = sMessageBuilt & "CompressionOff: " & objItem.CompressionOff
        sMessageBuilt = sMessageBuilt & "CompressionOn: " & objItem.CompressionOn
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "ConfigurationDialog: " & objItem.ConfigurationDialog
        sMessageBuilt = sMessageBuilt & "CountriesSupported: " & objItem.CountriesSupported
        sMessageBuilt = sMessageBuilt & "CountrySelected: " & objItem.CountrySelected
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "CurrentPasswords: " & objItem.CurrentPasswords
        sMessageBuilt = sMessageBuilt & "DCB: " & objItem.DCB
        sMessageBuilt = sMessageBuilt & "Default: " & objItem.Default
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "DeviceLoader: " & objItem.DeviceLoader
        sMessageBuilt = sMessageBuilt & "DeviceType: " & objItem.DeviceType
        sMessageBuilt = sMessageBuilt & "DialType: " & objItem.DialType
        sMessageBuilt = sMessageBuilt & "DriverDate: " & objItem.DriverDate
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorControlForced: " & objItem.ErrorControlForced
        sMessageBuilt = sMessageBuilt & "ErrorControlInfo: " & objItem.ErrorControlInfo
        sMessageBuilt = sMessageBuilt & "ErrorControlOff: " & objItem.ErrorControlOff
        sMessageBuilt = sMessageBuilt & "ErrorControlOn: " & objItem.ErrorControlOn
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "FlowControlHard: " & objItem.FlowControlHard
        sMessageBuilt = sMessageBuilt & "FlowControlOff: " & objItem.FlowControlOff
        sMessageBuilt = sMessageBuilt & "FlowControlSoft: " & objItem.FlowControlSoft
        sMessageBuilt = sMessageBuilt & "InactivityScale: " & objItem.InactivityScale
        sMessageBuilt = sMessageBuilt & "InactivityTimeout: " & objItem.InactivityTimeout
        sMessageBuilt = sMessageBuilt & "Index: " & objItem.Index
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "MaxBaudRateToPhone: " & objItem.MaxBaudRateToPhone
        sMessageBuilt = sMessageBuilt & "MaxBaudRateToSerialPort: " & objItem.MaxBaudRateToSerialPort
        sMessageBuilt = sMessageBuilt & "MaxNumberOfPasswords: " & objItem.MaxNumberOfPasswords
        sMessageBuilt = sMessageBuilt & "Model: " & objItem.Model
        sMessageBuilt = sMessageBuilt & "ModemInfPath: " & objItem.ModemInfPath
        sMessageBuilt = sMessageBuilt & "ModemInfSection: " & objItem.ModemInfSection
        sMessageBuilt = sMessageBuilt & "ModulationBell: " & objItem.ModulationBell
        sMessageBuilt = sMessageBuilt & "ModulationCCITT: " & objItem.ModulationCCITT
        sMessageBuilt = sMessageBuilt & "ModulationScheme: " & objItem.ModulationScheme
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PortSubClass: " & objItem.PortSubClass
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "Prefix: " & objItem.Prefix
        sMessageBuilt = sMessageBuilt & "Properties: " & objItem.Properties
        sMessageBuilt = sMessageBuilt & "ProviderName: " & objItem.ProviderName
        sMessageBuilt = sMessageBuilt & "Pulse: " & objItem.Pulse
        sMessageBuilt = sMessageBuilt & "Reset: " & objItem.Reset
        sMessageBuilt = sMessageBuilt & "ResponsesKeyName: " & objItem.ResponsesKeyName
        sMessageBuilt = sMessageBuilt & "RingsBeforeAnswer: " & objItem.RingsBeforeAnswer
        sMessageBuilt = sMessageBuilt & "SpeakerModeDial: " & objItem.SpeakerModeDial
        sMessageBuilt = sMessageBuilt & "SpeakerModeOff: " & objItem.SpeakerModeOff
        sMessageBuilt = sMessageBuilt & "SpeakerModeOn: " & objItem.SpeakerModeOn
        sMessageBuilt = sMessageBuilt & "SpeakerModeSetup: " & objItem.SpeakerModeSetup
        sMessageBuilt = sMessageBuilt & "SpeakerVolumeHigh: " & objItem.SpeakerVolumeHigh
        sMessageBuilt = sMessageBuilt & "SpeakerVolumeInfo: " & objItem.SpeakerVolumeInfo
        sMessageBuilt = sMessageBuilt & "SpeakerVolumeLow: " & objItem.SpeakerVolumeLow
        sMessageBuilt = sMessageBuilt & "SpeakerVolumeMed: " & objItem.SpeakerVolumeMed
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "StringFormat: " & objItem.StringFormat
        sMessageBuilt = sMessageBuilt & "SupportsCallback: " & objItem.SupportsCallback
        sMessageBuilt = sMessageBuilt & "SupportsSynchronousConnect: " & objItem.SupportsSynchronousConnect
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "Terminator: " & objItem.Terminator
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset
        sMessageBuilt = sMessageBuilt & "Tone: " & objItem.Tone
        sMessageBuilt = sMessageBuilt & "VoiceSwitchFeature: " & objItem.VoiceSwitchFeature
        
    

End Sub


Sub Win32_PowerManagementEvent(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PowerManagementEvent", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "EventType: " & objItem.EventType
        sMessageBuilt = sMessageBuilt & "OEMEventCode: " & objItem.OEMEventCode
        
    

End Sub


Sub Win32_Printer(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_Printer", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Attributes: " & objItem.Attributes
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "AveragePagesPerMinute: " & objItem.AveragePagesPerMinute
        sMessageBuilt = sMessageBuilt & "Capabilities: " & objItem.Capabilities
        sMessageBuilt = sMessageBuilt & "CapabilityDescriptions: " & objItem.CapabilityDescriptions
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "DefaultPriority: " & objItem.DefaultPriority
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DetectedErrorState: " & objItem.DetectedErrorState
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "DriverName: " & objItem.DriverName
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "HorizontalResolution: " & objItem.HorizontalResolution
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "JobCountSinceLastReset: " & objItem.JobCountSinceLastReset
        sMessageBuilt = sMessageBuilt & "LanguagesSupported: " & objItem.LanguagesSupported
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "Location: " & objItem.Location
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "PaperSizesSupported: " & objItem.PaperSizesSupported
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PortName: " & objItem.PortName
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "PrinterPaperNames: " & objItem.PrinterPaperNames
        sMessageBuilt = sMessageBuilt & "PrinterState: " & objItem.PrinterState
        sMessageBuilt = sMessageBuilt & "PrinterStatus: " & objItem.PrinterStatus
        sMessageBuilt = sMessageBuilt & "PrintJobDataType: " & objItem.PrintJobDataType
        sMessageBuilt = sMessageBuilt & "PrintProcessor: " & objItem.PrintProcessor
        sMessageBuilt = sMessageBuilt & "SeparatorFile: " & objItem.SeparatorFile
        sMessageBuilt = sMessageBuilt & "ServerName: " & objItem.ServerName
        sMessageBuilt = sMessageBuilt & "ShareName: " & objItem.ShareName
        sMessageBuilt = sMessageBuilt & "SpoolEnabled: " & objItem.SpoolEnabled
        sMessageBuilt = sMessageBuilt & "StartTime: " & objItem.StartTime
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset
        sMessageBuilt = sMessageBuilt & "UntilTime: " & objItem.UntilTime
        sMessageBuilt = sMessageBuilt & "VerticalResolution: " & objItem.VerticalResolution
        
    

End Sub


Sub Win32_PrinterConfiguration(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PrinterConfiguration", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "BitsPerPel: " & objItem.BitsPerPel
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "Collate: " & objItem.Collate
        sMessageBuilt = sMessageBuilt & "Color: " & objItem.Color
        sMessageBuilt = sMessageBuilt & "Copies: " & objItem.Copies
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceName: " & objItem.DeviceName
        sMessageBuilt = sMessageBuilt & "DisplayFlags: " & objItem.DisplayFlags
        sMessageBuilt = sMessageBuilt & "DisplayFrequency: " & objItem.DisplayFrequency
        sMessageBuilt = sMessageBuilt & "DitherType: " & objItem.DitherType
        sMessageBuilt = sMessageBuilt & "DriverVersion: " & objItem.DriverVersion
        sMessageBuilt = sMessageBuilt & "Duplex: " & objItem.Duplex
        sMessageBuilt = sMessageBuilt & "FormName: " & objItem.FormName
        sMessageBuilt = sMessageBuilt & "HorizontalResolution: " & objItem.HorizontalResolution
        sMessageBuilt = sMessageBuilt & "ICMIntent: " & objItem.ICMIntent
        sMessageBuilt = sMessageBuilt & "ICMMethod: " & objItem.ICMMethod
        sMessageBuilt = sMessageBuilt & "LogPixels: " & objItem.LogPixels
        sMessageBuilt = sMessageBuilt & "MediaType: " & objItem.MediaType
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "Orientation: " & objItem.Orientation
        sMessageBuilt = sMessageBuilt & "PaperLength: " & objItem.PaperLength
        sMessageBuilt = sMessageBuilt & "PaperSize: " & objItem.PaperSize
        sMessageBuilt = sMessageBuilt & "PaperWidth: " & objItem.PaperWidth
        sMessageBuilt = sMessageBuilt & "PelsHeight: " & objItem.PelsHeight
        sMessageBuilt = sMessageBuilt & "PelsWidth: " & objItem.PelsWidth
        sMessageBuilt = sMessageBuilt & "PrintQuality: " & objItem.PrintQuality
        sMessageBuilt = sMessageBuilt & "Scale: " & objItem.Scale
        sMessageBuilt = sMessageBuilt & "SettingID: " & objItem.SettingID
        sMessageBuilt = sMessageBuilt & "SpecificationVersion: " & objItem.SpecificationVersion
        sMessageBuilt = sMessageBuilt & "TTOption: " & objItem.TTOption
        sMessageBuilt = sMessageBuilt & "VerticalResolution: " & objItem.VerticalResolution
        sMessageBuilt = sMessageBuilt & "XResolution: " & objItem.XResolution
        sMessageBuilt = sMessageBuilt & "YResolution: " & objItem.YResolution
        
    

End Sub


Sub Win32_PrintJob(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_PrintJob", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "DataType: " & objItem.DataType
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "Document: " & objItem.Document
        sMessageBuilt = sMessageBuilt & "DriverName: " & objItem.DriverName
        sMessageBuilt = sMessageBuilt & "ElapsedTime: " & objItem.ElapsedTime
        sMessageBuilt = sMessageBuilt & "HostPrintQueue: " & objItem.HostPrintQueue
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "JobId: " & objItem.JobId
        sMessageBuilt = sMessageBuilt & "JobStatus: " & objItem.JobStatus
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "Notify: " & objItem.Notify
        sMessageBuilt = sMessageBuilt & "Owner: " & objItem.Owner
        sMessageBuilt = sMessageBuilt & "PagesPrinted: " & objItem.PagesPrinted
        sMessageBuilt = sMessageBuilt & "Parameters: " & objItem.Parameters
        sMessageBuilt = sMessageBuilt & "PrintProcessor: " & objItem.PrintProcessor
        sMessageBuilt = sMessageBuilt & "Priority: " & objItem.Priority
        sMessageBuilt = sMessageBuilt & "Size: " & objItem.Size
        sMessageBuilt = sMessageBuilt & "StartTime: " & objItem.StartTime
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusMask: " & objItem.StatusMask
        sMessageBuilt = sMessageBuilt & "TimeSubmitted: " & objItem.TimeSubmitted
        sMessageBuilt = sMessageBuilt & "TotalPages: " & objItem.TotalPages
        sMessageBuilt = sMessageBuilt & "UntilTime: " & objItem.UntilTime
        
    

End Sub


Sub Win32_SerialPort(objItem)
    

On Error Resume Next
    
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "Availability: " & objItem.Availability
        sMessageBuilt = sMessageBuilt & "Binary: " & objItem.Binary
        sMessageBuilt = sMessageBuilt & "Capabilities: " & objItem.Capabilities
        sMessageBuilt = sMessageBuilt & "CapabilityDescriptions: " & objItem.CapabilityDescriptions
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
        sMessageBuilt = sMessageBuilt & "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
        sMessageBuilt = sMessageBuilt & "CreationClassName: " & objItem.CreationClassName
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DeviceID: " & objItem.DeviceID
        sMessageBuilt = sMessageBuilt & "ErrorCleared: " & objItem.ErrorCleared
        sMessageBuilt = sMessageBuilt & "ErrorDescription: " & objItem.ErrorDescription
        sMessageBuilt = sMessageBuilt & "InstallDate: " & objItem.InstallDate
        sMessageBuilt = sMessageBuilt & "LastErrorCode: " & objItem.LastErrorCode
        sMessageBuilt = sMessageBuilt & "MaxBaudRate: " & objItem.MaxBaudRate
        sMessageBuilt = sMessageBuilt & "MaximumInputBufferSize: " & objItem.MaximumInputBufferSize
        sMessageBuilt = sMessageBuilt & "MaximumOutputBufferSize: " & objItem.MaximumOutputBufferSize
        sMessageBuilt = sMessageBuilt & "MaxNumberControlled: " & objItem.MaxNumberControlled
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "OSAutoDiscovered: " & objItem.OSAutoDiscovered
        sMessageBuilt = sMessageBuilt & "PNPDeviceID: " & objItem.PNPDeviceID
        sMessageBuilt = sMessageBuilt & "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
        sMessageBuilt = sMessageBuilt & "PowerManagementSupported: " & objItem.PowerManagementSupported
        sMessageBuilt = sMessageBuilt & "ProtocolSupported: " & objItem.ProtocolSupported
        sMessageBuilt = sMessageBuilt & "ProviderType: " & objItem.ProviderType
        sMessageBuilt = sMessageBuilt & "SettableBaudRate: " & objItem.SettableBaudRate
        sMessageBuilt = sMessageBuilt & "SettableDataBits: " & objItem.SettableDataBits
        sMessageBuilt = sMessageBuilt & "SettableFlowControl: " & objItem.SettableFlowControl
        sMessageBuilt = sMessageBuilt & "SettableParity: " & objItem.SettableParity
        sMessageBuilt = sMessageBuilt & "SettableParityCheck: " & objItem.SettableParityCheck
        sMessageBuilt = sMessageBuilt & "SettableRLSD: " & objItem.SettableRLSD
        sMessageBuilt = sMessageBuilt & "SettableStopBits: " & objItem.SettableStopBits
        sMessageBuilt = sMessageBuilt & "Status: " & objItem.Status
        sMessageBuilt = sMessageBuilt & "StatusInfo: " & objItem.StatusInfo
        sMessageBuilt = sMessageBuilt & "Supports16BitMode: " & objItem.Supports16BitMode
        sMessageBuilt = sMessageBuilt & "SupportsDTRDSR: " & objItem.SupportsDTRDSR
        sMessageBuilt = sMessageBuilt & "SupportsElapsedTimeouts: " & objItem.SupportsElapsedTimeouts
        sMessageBuilt = sMessageBuilt & "SupportsIntTimeouts: " & objItem.SupportsIntTimeouts
        sMessageBuilt = sMessageBuilt & "SupportsParityCheck: " & objItem.SupportsParityCheck
        sMessageBuilt = sMessageBuilt & "SupportsRLSD: " & objItem.SupportsRLSD
        sMessageBuilt = sMessageBuilt & "SupportsRTSCTS: " & objItem.SupportsRTSCTS
        sMessageBuilt = sMessageBuilt & "SupportsSpecialCharacters: " & objItem.SupportsSpecialCharacters
        sMessageBuilt = sMessageBuilt & "SupportsXOnXOff: " & objItem.SupportsXOnXOff
        sMessageBuilt = sMessageBuilt & "SupportsXOnXOffSet: " & objItem.SupportsXOnXOffSet
        sMessageBuilt = sMessageBuilt & "SystemCreationClassName: " & objItem.SystemCreationClassName
        sMessageBuilt = sMessageBuilt & "SystemName: " & objItem.SystemName
        sMessageBuilt = sMessageBuilt & "TimeOfLastReset: " & objItem.TimeOfLastReset
        
    

End Sub


Sub Win32_SerialPortConfiguration(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPortConfiguration", , 48)
    
        
        sMessageBuilt = sMessageBuilt & "AbortReadWriteOnError: " & objItem.AbortReadWriteOnError
        sMessageBuilt = sMessageBuilt & "BaudRate: " & objItem.BaudRate
        sMessageBuilt = sMessageBuilt & "BinaryModeEnabled: " & objItem.BinaryModeEnabled
        sMessageBuilt = sMessageBuilt & "BitsPerByte: " & objItem.BitsPerByte
        sMessageBuilt = sMessageBuilt & "Caption: " & objItem.Caption
        sMessageBuilt = sMessageBuilt & "ContinueXMitOnXOff: " & objItem.ContinueXMitOnXOff
        sMessageBuilt = sMessageBuilt & "CTSOutflowControl: " & objItem.CTSOutflowControl
        sMessageBuilt = sMessageBuilt & "Description: " & objItem.Description
        sMessageBuilt = sMessageBuilt & "DiscardNULLBytes: " & objItem.DiscardNULLBytes
        sMessageBuilt = sMessageBuilt & "DSROutflowControl: " & objItem.DSROutflowControl
        sMessageBuilt = sMessageBuilt & "DSRSensitivity: " & objItem.DSRSensitivity
        sMessageBuilt = sMessageBuilt & "DTRFlowControlType: " & objItem.DTRFlowControlType
        sMessageBuilt = sMessageBuilt & "EOFCharacter: " & objItem.EOFCharacter
        sMessageBuilt = sMessageBuilt & "ErrorReplaceCharacter: " & objItem.ErrorReplaceCharacter
        sMessageBuilt = sMessageBuilt & "ErrorReplacementEnabled: " & objItem.ErrorReplacementEnabled
        sMessageBuilt = sMessageBuilt & "EventCharacter: " & objItem.EventCharacter
        sMessageBuilt = sMessageBuilt & "IsBusy: " & objItem.IsBusy
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & "Parity: " & objItem.Parity
        sMessageBuilt = sMessageBuilt & "ParityCheckEnabled: " & objItem.ParityCheckEnabled
        sMessageBuilt = sMessageBuilt & "RTSFlowControlType: " & objItem.RTSFlowControlType
        sMessageBuilt = sMessageBuilt & "SettingID: " & objItem.SettingID
        sMessageBuilt = sMessageBuilt & "StopBits: " & objItem.StopBits
        sMessageBuilt = sMessageBuilt & "XOffCharacter: " & objItem.XOffCharacter
        sMessageBuilt = sMessageBuilt & "XOffXMitThreshold: " & objItem.XOffXMitThreshold
        sMessageBuilt = sMessageBuilt & "XOnCharacter: " & objItem.XOnCharacter
        sMessageBuilt = sMessageBuilt & "XOnXMitThreshold: " & objItem.XOnXMitThreshold
        sMessageBuilt = sMessageBuilt & "XOnXOffInFlowControl: " & objItem.XOnXOffInFlowControl
        sMessageBuilt = sMessageBuilt & "XOnXOffOutFlowControl: " & objItem.XOnXOffOutFlowControl
        
    

End Sub

Sub Win32_Service(objItem)

On Error Resume Next
    
    'Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPortConfiguration", , 48)
            
        sMessageBuilt = sMessageBuilt & "Started: " & objItem.Started
        sMessageBuilt = sMessageBuilt & ";"
        sMessageBuilt = sMessageBuilt & "State: " & objItem.State
        sMessageBuilt = sMessageBuilt & ";"
        sMessageBuilt = sMessageBuilt & "Name: " & objItem.Name
        sMessageBuilt = sMessageBuilt & ";"
        sMessageBuilt = sMessageBuilt & "DisplayName: " & objItem.DisplayName
        sMessageBuilt = sMessageBuilt & ";&&&"

End Sub




Private Sub Class_Initialize()
    WMIHardwareItemCount
End Sub
