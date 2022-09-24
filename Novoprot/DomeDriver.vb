'tabs=4
' --------------------------------------------------------------------------------
' TODO fill in this information for your driver, then remove this line!
'
' ASCOM Dome driver for Romanofafard
'
' Description:	Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam 
'				nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam 
'				erat, sed diam voluptua. At vero eos et accusam et justo duo 
'				dolores et ea rebum. Stet clita kasd gubergren, no sea takimata 
'				sanctus est Lorem ipsum dolor sit amet.
'
' Implements:	ASCOM Dome interface version: 1.0
' Author:		(XXX) Your N. Here <your@email.here>
'
' Edit Log:
'
' Date			Who	Vers	Description
' -----------	---	-----	-------------------------------------------------------
' dd-mmm-yyyy	XXX	1.0.0	Initial edit, from Dome template
' ---------------------------------------------------------------------------------
'
'
' Your driver's ID is ASCOM.Romanofafard.Dome
'
' The Guid attribute sets the CLSID for ASCOM.DeviceName.Dome
' The ClassInterface/None attribute prevents an empty interface called
' _Dome from being created and used as the [default] interface
'

' This definition is used to select code that's only applicable for one device type
#Const Device = "Dome"

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports ASCOM
Imports ASCOM.Astrometry
Imports ASCOM.Astrometry.AstroUtils
Imports ASCOM.DeviceInterface
Imports ASCOM.Utilities

<Guid("02dff19d-1558-46a4-bb7e-004962f3b523")>
<ClassInterface(ClassInterfaceType.None)>
Public Class Dome

    ' The Guid attribute sets the CLSID for ASCOM.Romanofafard.Dome
    ' The ClassInterface/None attribute prevents an empty interface called
    ' _Romanofafard from being created and used as the [default] interface

    ' TODO Replace the not implemented exceptions with code to implement the function or
    ' throw the appropriate ASCOM exception.
    '
    Implements IDomeV2

    '
    ' Driver ID and descriptive string that shows in the Chooser
    '
    Friend Shared driverID As String = "ASCOM.Romanofafard.Dome"
    Private Shared driverDescription As String = "Romanofafard Dome"

    Private objSerial As Serial

    Friend Shared comPortProfileName As String = "COM Port" 'Constants used for Profile persistence
    Friend Shared traceStateProfileName As String = "Trace Level"
    Friend Shared comPortDefault As String = "COM1"
    Friend Shared traceStateDefault As String = "False"

    Friend Shared comPort As String ' Variables to hold the current device configuration
    Friend Shared traceState As Boolean

    Private connectedState As Boolean ' Private variable to hold the connected state
    Private utilities As Util ' Private variable to hold an ASCOM Utilities object
    Private astroUtilities As AstroUtils ' Private variable to hold an AstroUtils object to provide the Range method
    Private TL As TraceLogger ' Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)

    '
    ' Constructor - Must be public for COM registration!
    '
    Public Sub New()

        ReadProfile() ' Read device configuration from the ASCOM Profile store
        TL = New TraceLogger("", "Romanofafard")
        TL.Enabled = traceState
        TL.LogMessage("Dome", "Starting initialisation")

        connectedState = False ' Initialise connected to false
        utilities = New Util() ' Initialise util object
        astroUtilities = New AstroUtils 'Initialise new astro utilities object

        'TODO: Implement your additional construction here

        TL.LogMessage("Dome", "Completed initialisation")
    End Sub

    '
    ' PUBLIC COM INTERFACE IDomeV2 IMPLEMENTATION
    '

#Region "Common properties and methods"
    ''' <summary>
    ''' Displays the Setup Dialog form.
    ''' If the user clicks the OK button to dismiss the form, then
    ''' the new settings are saved, otherwise the old values are reloaded.
    ''' THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
    ''' </summary>
    Public Sub SetupDialog() Implements IDomeV2.SetupDialog
        ' consider only showing the setup dialog if not connected
        ' or call a different dialog if connected
        If IsConnected Then
            System.Windows.Forms.MessageBox.Show("Already connected, just press OK")
        End If

        Using F As SetupDialogForm = New SetupDialogForm()
            Dim result As System.Windows.Forms.DialogResult = F.ShowDialog()
            If result = DialogResult.OK Then
                WriteProfile() ' Persist device configuration values to the ASCOM Profile store
            End If
        End Using
    End Sub

    ''' <summary>Returns the list of custom action names supported by this driver.</summary>
    ''' <value>An ArrayList of strings (SafeArray collection) containing the names of supported actions.</value>
    Public ReadOnly Property SupportedActions() As ArrayList Implements IDomeV2.SupportedActions
        Get
            TL.LogMessage("SupportedActions Get", "Returning empty arraylist")
            Return New ArrayList()
        End Get
    End Property

    ''' <summary>Invokes the specified device-specific custom action.</summary>
    ''' <param name="ActionName">A well known name agreed by interested parties that represents the action to be carried out.</param>
    ''' <param name="ActionParameters">List of required parameters or an <see cref="String.Empty">Empty String</see> if none are required.</param>
    ''' <returns>A string response. The meaning of returned strings is set by the driver author.
    ''' <para>Suppose filter wheels start to appear with automatic wheel changers; new actions could be <c>QueryWheels</c> and <c>SelectWheel</c>. The former returning a formatted list
    ''' of wheel names and the second taking a wheel name and making the change, returning appropriate values to indicate success or failure.</para>
    ''' </returns>
    Public Function Action(ByVal ActionName As String, ByVal ActionParameters As String) As String Implements IDomeV2.Action
        Throw New ActionNotImplementedException("Action " & ActionName & " is not supported by this driver")
    End Function

    ''' <summary>
    ''' Transmits an arbitrary string to the device and does not wait for a response.
    ''' Optionally, protocol framing characters may be added to the string before transmission.
    ''' </summary>
    ''' <param name="Command">The literal command string to be transmitted.</param>
    ''' <param name="Raw">
    ''' if set to <c>True</c> the string is transmitted 'as-is'.
    ''' If set to <c>False</c> then protocol framing characters may be added prior to transmission.
    ''' </param>
    Public Sub CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False) Implements IDomeV2.CommandBlind
        CheckConnected("CommandBlind")
        ' TODO The optional CommandBlind method should either be implemented OR throw a MethodNotImplementedException
        ' If implemented, CommandBlind must send the supplied command to the mount And return immediately without waiting for a response

        Throw New MethodNotImplementedException("CommandBlind")
    End Sub

    ''' <summary>
    ''' Transmits an arbitrary string to the device and waits for a boolean response.
    ''' Optionally, protocol framing characters may be added to the string before transmission.
    ''' </summary>
    ''' <param name="Command">The literal command string to be transmitted.</param>
    ''' <param name="Raw">
    ''' if set to <c>True</c> the string is transmitted 'as-is'.
    ''' If set to <c>False</c> then protocol framing characters may be added prior to transmission.
    ''' </param>
    ''' <returns>
    ''' Returns the interpreted boolean response received from the device.
    ''' </returns>
    Public Function CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean _
        Implements IDomeV2.CommandBool
        CheckConnected("CommandBool")
        ' TODO The optional CommandBool method should either be implemented OR throw a MethodNotImplementedException
        ' If implemented, CommandBool must send the supplied command to the mount, wait for a response and parse this to return a True Or False value

        ' Dim retString as String = CommandString(command, raw) ' Send the command And wait for the response
        ' Dim retBool as Boolean = XXXXXXXXXXXXX ' Parse the returned string And create a boolean True / False value
        ' Return retBool ' Return the boolean value to the client

        Throw New MethodNotImplementedException("CommandBool")
    End Function

    ''' <summary>
    ''' Transmits an arbitrary string to the device and waits for a string response.
    ''' Optionally, protocol framing characters may be added to the string before transmission.
    ''' </summary>
    ''' <param name="Command">The literal command string to be transmitted.</param>
    ''' <param name="Raw">
    ''' if set to <c>True</c> the string is transmitted 'as-is'.
    ''' If set to <c>False</c> then protocol framing characters may be added prior to transmission.
    ''' </param>
    ''' <returns>
    ''' Returns the string response received from the device.
    ''' </returns>
    Public Function CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String _
        Implements IDomeV2.CommandString
        CheckConnected("CommandString")
        ' TODO The optional CommandString method should either be implemented OR throw a MethodNotImplementedException
        ' If implemented, CommandString must send the supplied command to the mount and wait for a response before returning this to the client

        Throw New MethodNotImplementedException("CommandString")
    End Function

    ''' <summary>
    ''' Set True to connect to the device hardware. Set False to disconnect from the device hardware.
    ''' You can also read the property to check whether it is connected. This reports the current hardware state.
    ''' </summary>
    ''' <value><c>True</c> if connected to the hardware; otherwise, <c>False</c>.</value>
    Public Property Connected() As Boolean Implements IDomeV2.Connected
        Get
            TL.LogMessage("Connected Get", IsConnected.ToString())
            Return IsConnected
        End Get
        Set(value As Boolean)
            TL.LogMessage("Connected Set", value.ToString())
            If value = IsConnected Then
                Return
            End If

            If value Then
                objSerial = New ASCOM.Utilities.Serial
                objSerial.Port = 3
                objSerial.Speed = 57600
                objSerial.Connected = True
                connectedState = True
                TL.LogMessage("Connected Set", "Connecting to port " + comPort)
                ' TODO connect to the device
            Else
                objSerial.Connected = False
                connectedState = False
                TL.LogMessage("Connected Set", "Disconnecting from port " + comPort)
                Dispose()
                ' TODO disconnect from the device
            End If
        End Set
    End Property

    ''' <summary>
    ''' Returns a description of the device, such as manufacturer and modelnumber. Any ASCII characters may be used.
    ''' </summary>
    ''' <value>The description.</value>
    Public ReadOnly Property Description As String Implements IDomeV2.Description
        Get
            ' this pattern seems to be needed to allow a public property to return a private field
            Dim d As String = driverDescription
            TL.LogMessage("Description Get", d)
            Return d
        End Get
    End Property

    ''' <summary>
    ''' Descriptive and version information about this ASCOM driver.
    ''' </summary>
    Public ReadOnly Property DriverInfo As String Implements IDomeV2.DriverInfo
        Get
            Dim m_version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            ' TODO customise this driver description
            Dim s_driverInfo As String = "Information about the driver itself. Version: " + m_version.Major.ToString() + "." + m_version.Minor.ToString()
            TL.LogMessage("DriverInfo Get", s_driverInfo)
            Return s_driverInfo
        End Get
    End Property

    ''' <summary>
    ''' A string containing only the major and minor version of the driver formatted as 'm.n'.
    ''' </summary>
    Public ReadOnly Property DriverVersion() As String Implements IDomeV2.DriverVersion
        Get
            ' Get our own assembly and report its version number
            TL.LogMessage("DriverVersion Get", Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2))
            Return Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2)
        End Get
    End Property

    ''' <summary>
    ''' The interface version number that this device supports. 
    ''' </summary>
    Public ReadOnly Property InterfaceVersion() As Short Implements IDomeV2.InterfaceVersion
        Get
            TL.LogMessage("InterfaceVersion Get", "2")
            Return 2
        End Get
    End Property

    ''' <summary>
    ''' The short name of the driver, for display purposes
    ''' </summary>
    Public ReadOnly Property Name As String Implements IDomeV2.Name
        Get
            Dim s_name As String = "Romanofafard"
            TL.LogMessage("Name Get", s_name)
            Return s_name
        End Get
    End Property

    ''' <summary>
    ''' Dispose the late-bound interface, if needed. Will release it via COM
    ''' if it is a COM object, else if native .NET will just dereference it
    ''' for GC.
    ''' </summary>
    Public Sub Dispose() Implements IDomeV2.Dispose
        ' Clean up the trace logger and util objects
        TL.Enabled = False
        TL.Dispose()
        TL = Nothing
        utilities.Dispose()
        utilities = Nothing
        astroUtilities.Dispose()
        astroUtilities = Nothing
    End Sub

#End Region

#Region "IDome Implementation"

    Private domeShutterState As Boolean = False ' Variable to hold the open/closed status of the shutter, true = Open

    ''' <summary>
    ''' Immediately stops any and all movement of the dome.
    ''' </summary>
    Public Sub AbortSlew() Implements IDomeV2.AbortSlew
        ' This is a mandatory parameter but we have no action to take in this simple driver
        TL.LogMessage("AbortSlew", "Completed")
    End Sub

    ''' <summary>
    ''' The altitude (degrees, horizon zero and increasing positive to 90 zenith) of the part of the sky that the observer wishes to observe.
    ''' </summary>
    Public ReadOnly Property Altitude() As Double Implements IDomeV2.Altitude
        Get
            TL.LogMessage("Altitude Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("Altitude", False)
        End Get
    End Property

    ''' <summary>
    ''' <para><see langword="True" /> when the dome is in the home position. Raises an error if not supported.</para>
    ''' <para>
    ''' This is normally used following a <see cref="FindHome" /> operation. The value is reset
    ''' with any azimuth slew operation that moves the dome away from the home position.
    ''' </para>
    ''' <para>
    ''' <see cref="AtHome" /> may optionally also become True during normal slew operations, if the
    ''' dome passes through the home position and the dome controller hardware is capable of
    ''' detecting that; or at the end of a slew operation if the dome comes to rest at the home
    ''' position.
    ''' </para>
    ''' </summary>
    Public ReadOnly Property AtHome() As Boolean Implements IDomeV2.AtHome
        Get
            objSerial.Transmit("ATHOME#")
            Dim s As String
            s = objSerial.ReceiveTerminated("#")
            s = s.Replace("#", "")
            Return CShort(s)
        End Get
    End Property

    ''' <summary>
    ''' <see langword="True" /> if the dome is in the programmed park position.
    ''' </summary>
    Public ReadOnly Property AtPark() As Boolean Implements IDomeV2.AtPark
        Get
            TL.LogMessage("AtPark", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("AtPark", False)
        End Get
    End Property

    ''' <summary>
    ''' The dome azimuth (degrees, North zero and increasing clockwise, i.e., 90 East, 180 South, 270 West). North is true north and not magnetic north.
    ''' </summary>
    Public ReadOnly Property Azimuth() As Double Implements IDomeV2.Azimuth
        Get
            objSerial.Transmit("AZIMUTH#")
            Dim s As String
            s = objSerial.ReceiveTerminated("#")
            s = s.Replace("#", "")
            Return CShort(s)
        End Get
    End Property

    ''' <summary>
    ''' <see langword="True" /> if the driver can perform a search for home position.
    ''' </summary>
    Public ReadOnly Property CanFindHome() As Boolean Implements IDomeV2.CanFindHome
        Get
            TL.LogMessage("CanFindHome Get", True.ToString())
            Return True
        End Get
    End Property

    ''' <summary>
    ''' <see langword="True" /> if the driver is capable of parking the dome.
    ''' </summary>
    Public ReadOnly Property CanPark() As Boolean Implements IDomeV2.CanPark
        Get
            TL.LogMessage("CanPark Get", False.ToString())
            Return False
        End Get
    End Property

    ''' <summary>
    ''' <see langword="True" /> if driver is capable of setting dome altitude.
    ''' </summary>
    Public ReadOnly Property CanSetAltitude() As Boolean Implements IDomeV2.CanSetAltitude
        Get
            TL.LogMessage("CanSetAltitude Get", False.ToString())
            Return False
        End Get
    End Property

    ''' <summary>
    ''' <see langword="True" /> if driver is capable of rotating the dome. Must be <see "langword="False" /> for a 
    ''' roll-off roof or clamshell.
    ''' </summary>
    Public ReadOnly Property CanSetAzimuth() As Boolean Implements IDomeV2.CanSetAzimuth
        Get
            TL.LogMessage("CanSetAzimuth Get", False.ToString())
            Return True
        End Get
    End Property

    ''' <summary>
    ''' <see langword="True" /> if the driver can set the dome park position.
    ''' </summary>
    Public ReadOnly Property CanSetPark() As Boolean Implements IDomeV2.CanSetPark
        Get
            TL.LogMessage("CanSetPark Get", False.ToString())
            Return False
        End Get
    End Property

    ''' <summary>
    ''' <see langword="True" /> if the driver is capable of opening and closing the shutter or roof
    ''' mechanism.
    ''' </summary>
    Public ReadOnly Property CanSetShutter() As Boolean Implements IDomeV2.CanSetShutter
        Get
            TL.LogMessage("CanSetShutter Get", True.ToString())
            Return True
        End Get
    End Property

    ''' <summary>
    ''' <see langword="true" /> if the dome hardware supports slaving to a telescope.
    ''' </summary>
    Public ReadOnly Property CanSlave() As Boolean Implements IDomeV2.CanSlave
        Get
            TL.LogMessage("CanSlave Get", False.ToString())
            Return False
        End Get
    End Property

    ''' <summary>
    ''' <see langword="true" /> if the driver is capable of synchronizing the dome azimuth position
    ''' using the <see cref="SyncToAzimuth" /> method.
    ''' </summary>
    Public ReadOnly Property CanSyncAzimuth() As Boolean Implements IDomeV2.CanSyncAzimuth
        Get
            TL.LogMessage("CanSyncAzimuth Get", False.ToString())
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Close the shutter or otherwise shield the telescope from the sky.
    ''' </summary>
    Public Sub CloseShutter() Implements IDomeV2.CloseShutter
        TL.LogMessage("CloseShutter", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("CloseShutter")
    End Sub

    ''' <summary>
    ''' Start operation to search for the dome home position.
    ''' </summary>
    Public Sub FindHome() Implements IDomeV2.FindHome
        objSerial.Transmit("FINDHOME#")
    End Sub

    ''' <summary>
    ''' Open shutter or otherwise expose telescope to the sky.
    ''' </summary>
    Public Sub OpenShutter() Implements IDomeV2.OpenShutter
        TL.LogMessage("OpenShutter", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("OpenShutter")
    End Sub

    ''' <summary>
    ''' Rotate dome in azimuth to park position.
    ''' </summary>
    Public Sub Park() Implements IDomeV2.Park
        TL.LogMessage("Park", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("Park")
    End Sub

    ''' <summary>
    ''' Set the current azimuth position of dome to the park position.
    ''' </summary>
    Public Sub SetPark() Implements IDomeV2.SetPark
        TL.LogMessage("SetPark", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SetPark")
    End Sub

    ''' <summary>
    ''' Gets the status of the dome shutter or roof structure.
    ''' </summary>
    Public ReadOnly Property ShutterStatus() As ShutterState Implements IDomeV2.ShutterStatus
        Get
            TL.LogMessage("CanSyncAzimuth Get", False.ToString())
            If (domeShutterState) Then
                TL.LogMessage("ShutterStatus", ShutterState.shutterOpen.ToString())
                Return ShutterState.shutterOpen
            Else
                TL.LogMessage("ShutterStatus", ShutterState.shutterClosed.ToString())
                Return ShutterState.shutterClosed
            End If
        End Get
    End Property

    ''' <summary>
    ''' <see langword="True"/> if the dome is slaved to the telescope in its hardware, else <see langword="False"/>.
    ''' </summary>
    Public Property Slaved() As Boolean Implements IDomeV2.Slaved
        Get
            TL.LogMessage("Slaved Get", False.ToString())
            Return False
        End Get
        Set(value As Boolean)
            TL.LogMessage("Slaved Set", "not implemented")
            Throw New ASCOM.PropertyNotImplementedException("Slaved", True)
        End Set
    End Property

    ''' <summary>
    ''' Ensure that the requested viewing altitude is available for observing.
    ''' </summary>
    ''' <param name="Altitude">
    ''' The desired viewing altitude (degrees, horizon zero and increasing positive to 90 degrees at the zenith)
    ''' </param>
    Public Sub SlewToAltitude(Altitude As Double) Implements IDomeV2.SlewToAltitude
        TL.LogMessage("SlewToAltitude", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToAltitude")
    End Sub

    ''' <summary>
    ''' Ensure that the requested viewing azimuth is available for observing.
    ''' The method should not block and the slew operation should complete asynchronously.
    ''' </summary>
    ''' <param name="Azimuth">
    ''' Desired viewing azimuth (degrees, North zero and increasing clockwise. i.e., 90 East,
    ''' 180 South, 270 West)
    ''' </param>
    Public Sub SlewToAzimuth(Azimuth As Double) Implements IDomeV2.SlewToAzimuth
        Dim s As String = Azimuth.ToString + "GOTO#"
        objSerial.Transmit(s)
    End Sub

    ''' <summary>
    ''' <see langword="True" /> if any part of the dome is currently moving or a move command has been issued, 
    ''' but the dome has not yet started to move. <see langword="False" /> if all dome components are stationary
    ''' and no move command has been issued. /> 
    ''' </summary>
    Public ReadOnly Property Slewing() As Boolean Implements IDomeV2.Slewing
        Get
            TL.LogMessage("Slewing Get", False.ToString())
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Synchronize the current position of the dome to the given azimuth.
    ''' </summary>
    ''' <param name="Azimuth">
    ''' Target azimuth (degrees, North zero and increasing clockwise. i.e., 90 East,
    ''' 180 South, 270 West)
    ''' </param>
    Public Sub SyncToAzimuth(Azimuth As Double) Implements IDomeV2.SyncToAzimuth
        TL.LogMessage("SyncToAzimuth", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SyncToAzimuth")
    End Sub

#End Region

#Region "Private properties and methods"
    ' here are some useful properties and methods that can be used as required
    ' to help with

#Region "ASCOM Registration"

    Private Shared Sub RegUnregASCOM(ByVal bRegister As Boolean)

        Using P As New Profile() With {.DeviceType = "Dome"}
            If bRegister Then
                P.Register(driverID, driverDescription)
            Else
                P.Unregister(driverID)
            End If
        End Using

    End Sub

    <ComRegisterFunction()>
    Public Shared Sub RegisterASCOM(ByVal T As Type)

        RegUnregASCOM(True)

    End Sub

    <ComUnregisterFunction()>
    Public Shared Sub UnregisterASCOM(ByVal T As Type)

        RegUnregASCOM(False)

    End Sub

#End Region

    ''' <summary>
    ''' Returns true if there is a valid connection to the driver hardware
    ''' </summary>
    Private ReadOnly Property IsConnected As Boolean
        Get
            ' TODO check that the driver hardware connection exists and is connected to the hardware
            If objSerial Is Nothing Then
                Return False
            Else

                If Not objSerial.Connected Then
                    Return False
                Else
                    Return True
                End If
            End If
        End Get
    End Property

    ''' <summary>
    ''' Use this function to throw an exception if we aren't connected to the hardware
    ''' </summary>
    ''' <param name="message"></param>
    Private Sub CheckConnected(ByVal message As String)
        If Not IsConnected Then
            Throw New NotConnectedException(message)
        End If
    End Sub

    ''' <summary>
    ''' Read the device configuration from the ASCOM Profile store
    ''' </summary>
    Friend Sub ReadProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "Dome"
            traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, String.Empty, traceStateDefault))
            comPort = driverProfile.GetValue(driverID, comPortProfileName, String.Empty, comPortDefault)
        End Using
    End Sub

    ''' <summary>
    ''' Write the device configuration to the  ASCOM  Profile store
    ''' </summary>
    Friend Sub WriteProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "Dome"
            driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString())
            If comPort IsNot Nothing Then
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString())
            End If
        End Using

    End Sub

#End Region

End Class
