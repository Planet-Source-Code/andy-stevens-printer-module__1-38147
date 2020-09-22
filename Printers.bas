Attribute VB_Name = "modPrinters"
Option Explicit

'*******************************************************************
'   Name        :   modPrinters
'   Purpose     :   Provides functions for manipulating system
'                   printers
'   Author      :   Andy Stevens
'   Date        :   August 2002
'*******************************************************************

'Declare API functions
Public Declare Function DeletePrinterConnection Lib "winspool.drv" Alias "DeletePrinterConnectionA" (ByVal pName As String) As Long

Public Function GetDefaultPrinter() As String

'*******************************************************************
'   Name        :   GetDefaultPrinter
'   Purpose     :   Retrieves the path of the current default printer
'                   on the local machine
'   Parameters  :   None
'   Returns     :   Default printer if successful, empty string if not
'   Author      :   Andy Stevens
'   Date        :   August 2002
'*******************************************************************

'Declare local variables
Dim strReturn   As String
Dim prnDefault  As Printer

'Declare local constants
Const FUNCTION_NAME As String = "modPrinters.GetDefaultPrinter"
    
On Error GoTo ErrorHandler
            
    'Set default return value
    strReturn = vbNullString
            
    'Get the current default printer
    Set prnDefault = Printer
    
    'Get the name of the default printer
    strReturn = prnDefault.DeviceName
       
CleanExit:

    'Kill the printer object
    If Not prnDefault Is Nothing Then
        Set prnDefault = Nothing
    End If

    GetDefaultPrinter = strReturn

    Exit Function

ErrorHandler:

    'Display the error
    MsgBox "Error No: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Has occured in " & FUNCTION_NAME & vbCrLf & _
           "Please contact Technical Support", vbOKOnly + vbCritical, App.Title
    
    Resume CleanExit
    
End Function

Public Function SetDefaultPrinter(ByVal v_strPrinterPath As String) As Boolean

'*******************************************************************
'   Name        :   SetDefaultPrinter
'   Purpose     :   Sets the default printer on the local machine
'   Parameters  :   v_strPrinterPath  : Path of printer to set as default
'   Returns     :   TRUE if successful, FALSE if not
'   Author      :   Andy Stevens
'   Date        :   August 2002
'*******************************************************************

'Declare local variables
Dim blnReturn   As Boolean
Dim objNetwork  As Object
    
'Declare local constants
Const FUNCTION_NAME As String = "modPrinters.SetDefaultPrinter"
    
On Error GoTo ErrorHandler

    'Set the default return value
    blnReturn = False
    
    'Create the Script Host object
    Set objNetwork = CreateObject("WScript.Network")
    
    'Set the printer to be the default
    objNetwork.SetDefaultPrinter v_strPrinterPath
    
    'Printer set
    blnReturn = True

CleanExit:

    'Kill the script host object
    If Not objNetwork Is Nothing Then
        Set objNetwork = Nothing
    End If

    SetDefaultPrinter = blnReturn

    Exit Function

ErrorHandler:

    'Display the error
    MsgBox "Error No: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Has occured in " & FUNCTION_NAME & vbCrLf & _
           "Please contact Technical Support", vbOKOnly + vbCritical, App.Title
    
    Resume CleanExit
    
End Function

Public Function AddPrinter(ByVal v_strPrinterPath As String, _
                           Optional ByVal v_strPrinterDriver = "") As Boolean

'*******************************************************************
'   Name        :   AddPrinter
'   Purpose     :   Adds the specified printer to the local machine
'   Parameters  :   v_strPrinterPath  : Path of printer to add
'                   v_strPrinterDriver: Driver of printer to add (Optional)
'   Returns     :   TRUE if printer added, FALSE if not
'   Author      :   Andy Stevens
'   Date        :   August 2002
'*******************************************************************

'Declare local variables
Dim blnReturn   As Boolean
Dim objNetwork  As Object
    
'Declare local constants
Const FUNCTION_NAME As String = "modPrinters.AddPrinter"
    
On Error GoTo ErrorHandler

    'Set the default return value
    blnReturn = False
    
    'Create the Script Host object
    Set objNetwork = CreateObject("WScript.Network")
    
    'Add the printer
    objNetwork.AddWindowsPrinterConnection v_strPrinterPath, v_strPrinterDriver
        
    'Printer added and set
    blnReturn = True

CleanExit:

    'Kill the script host object
    If Not objNetwork Is Nothing Then
        Set objNetwork = Nothing
    End If

    AddPrinter = blnReturn

    Exit Function

ErrorHandler:

    'Display the error
    MsgBox "Error No: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Has occured in " & FUNCTION_NAME & vbCrLf & _
           "Please contact Technical Support", vbOKOnly + vbCritical, App.Title
    
    Resume CleanExit
    
End Function

Public Function RemovePrinter(ByVal v_strPrinterPath As String) As Boolean

'*******************************************************************
'   Name        :   RemovePrinter
'   Purpose     :   Removes the specified printer from the local
'                   machine
'                   NB: API function used as RemovePrinterConnection
'                   method of Windows Script Host did not work!
'   Parameters  :   v_strPrinterPath   : Printer to remove
'   Returns     :   TRUE if printer removed, FALSE if not
'   Author      :   Andy Stevens
'   Date        :   August 2002
'*******************************************************************

'Declare local variables
Dim blnReturn   As Boolean
Dim lngReturn   As Long
    
'Declare local constants
Const FUNCTION_NAME As String = "modPrinters.RemovePrinter"
    
On Error GoTo ErrorHandler
        
    'Set the default return value
    blnReturn = False
        
    'Delete the printer
    lngReturn = DeletePrinterConnection(v_strPrinterPath)
    
    'Check the result
    If lngReturn <> 0 Then
    
        'Printer deleted
        blnReturn = True
        
    End If
    
CleanExit:

    RemovePrinter = blnReturn

    Exit Function

ErrorHandler:

    'Display the error
    MsgBox "Error No: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Has occured in " & FUNCTION_NAME & vbCrLf & _
           "Please contact Technical Support", vbOKOnly + vbCritical, App.Title
    
    Resume CleanExit
    
End Function

Public Function GetAllConnectedPrinters() As String()

'*******************************************************************
'   Name        :   GetAllConnectedPrinters
'   Purpose     :   Retrives all printers currently connected to
'                   the local machine
'   Parameters  :   None
'   Returns     :   String array containing printer names
'   Author      :   Andy Stevens
'   Date        :   August 2002
'*******************************************************************

'Declare local variables
Dim bytIndex    As Byte
Dim objNetwork  As Object
Dim objPrinters As Object
Dim bytCounter  As Byte
Dim strReturn() As String
    
'Declare local constants
Const FUNCTION_NAME As String = "modPrinters.GetAllConnectedPrinters"
    
On Error GoTo ErrorHandler

    'Set the index counter
    bytIndex = 0
    
    'Create the Script Host object
    Set objNetwork = CreateObject("WScript.Network")
    
    'Retrieve the printer details
    Set objPrinters = objNetwork.EnumPrinterConnections
    
    For bytCounter = 0 To objPrinters.Count - 1 Step 2
       
        'Re dimesion the array
        ReDim Preserve strReturn(bytIndex)
           
        'Populate the array
        strReturn(bytIndex) = objPrinters.Item(bytCounter + 1)
       
        'Increment the index
        bytIndex = bytIndex + 1
    
    Next bytCounter
    
CleanExit:

    'Kill the script host object
    If Not objNetwork Is Nothing Then
        Set objNetwork = Nothing
    End If

    'Kill the printers object
    If Not objPrinters Is Nothing Then
        Set objPrinters = Nothing
    End If

    GetAllConnectedPrinters = strReturn()

    Exit Function

ErrorHandler:

   'Display the error
    MsgBox "Error No: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Has occured in " & FUNCTION_NAME & vbCrLf & _
           "Please contact Technical Support", vbOKOnly + vbCritical, App.Title
    
    Resume CleanExit
    
End Function

Public Function IsPrinterConnected(ByVal v_strPrinterPath As String) As Boolean

'*******************************************************************
'   Name        :   IsPrinterConnected
'   Purpose     :   Checks to see if the specified printer is
'                   connected to the local machine
'   Parameters  :   v_strPrinterPath   : Printer to find
'   Returns     :   TRUE if printer connected, FALSE if not
'   Author      :   Andy Stevens
'   Date        :   August 2002
'*******************************************************************

'Declare local variables
Dim blnReturn   As Boolean
Dim objNetwork  As Object
Dim objPrinters As Object
Dim bytCounter  As Byte
    
'Declare local constants
Const FUNCTION_NAME As String = "modPrinters.IsPrinterConnected"
    
On Error GoTo ErrorHandler

    'Set the default return value
    blnReturn = False

    'Create the Script Host object
    Set objNetwork = CreateObject("WScript.Network")
    
    'Retrieve the printer details
    Set objPrinters = objNetwork.EnumPrinterConnections
    
    For bytCounter = 0 To objPrinters.Count - 1
       
        'Check for the specified printer
        If objPrinters.Item(bytCounter) = v_strPrinterPath Then
       
            'Printer Found
            blnReturn = True
            Exit For
            
        End If
       
    Next bytCounter
    
CleanExit:

    'Kill the script host object
    If Not objNetwork Is Nothing Then
        Set objNetwork = Nothing
    End If

    'Kill the printers object
    If Not objPrinters Is Nothing Then
        Set objPrinters = Nothing
    End If

    IsPrinterConnected = blnReturn

    Exit Function

ErrorHandler:

   'Display the error
    MsgBox "Error No: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Has occured in " & FUNCTION_NAME & vbCrLf & _
           "Please contact Technical Support", vbOKOnly + vbCritical, App.Title
    
    Resume CleanExit
    
End Function
