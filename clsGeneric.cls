VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As Dictionary

Private Sub Class_Initialize()
    ' Create a hash table to store the values
    Set mData = New Dictionary
    mData.CompareMode = TextCompare
End Sub


Private Sub Class_Terminate()
    ' Destroy the hash table
    Set mData = Nothing
End Sub


Public Function z_Dictionary() As Dictionary
    ' Return the hash table
    Set z_Dictionary = mData
End Function


Public Sub Inherit(ByRef obj As Variant)
    ' Inherit all the properties and methods of another object
    Dim dict As Dictionary
    Dim SuperClass As clsGeneric
    Dim i As Long
    Dim keys() As Variant
    
    If TypeName(obj) <> "clsGeneric" Then
        MsgBox "You can't inherit from this type of object!", vbCritical
        End
    End If
    
    Set SuperClass = obj
    Set dict = SuperClass.z_Dictionary
    keys = dict.keys
    For i = 0 To UBound(keys, 1)
        ' Inherit all except name
        If StrComp(keys(i), "_ClassName", vbTextCompare) <> 0 Then
            Change keys(i), dict(keys(i))
        End If
    Next i
End Sub


Private Sub Lookup(ByVal Name As String, ByRef Data As Variant)
    Dim dict As Dictionary
    Dim SuperClass As clsGeneric
    If mData.Exists(Name) Then
        ' Pull it from the object
        If IsObject(mData(Name)) Then
            Set Data = mData(Name)
        Else
            Data = mData(Name)
        End If
    ElseIf mData.Exists("_SuperClass") Then
        ' Pull it from the superclass
        Set SuperClass = mData("_SuperClass")
        Set dict = SuperClass.z_Dictionary
        If dict.Exists(Name) Then
            If IsObject(dict(Name)) Then
                Set Data = dict(Name)
            Else
                Data = dict(Name)
            End If
        Else
            Form1.Display "Class '" & dict("_ClassName") & "' doesn't have a property '" & Name & "'!"
        End If
    End If
End Sub


Private Sub Change(ByVal Name As String, ByRef Data As Variant)
    If mData.Exists(Name) Then
        mData.Remove Name
    End If
    mData.Add Name, Data
End Sub


Public Property Get Stat(ByVal Name As String) As Variant
    Lookup Name, Stat
End Property


Public Property Let Stat(ByVal Name As String, ByRef Value As Variant)
    Change Name, Value
End Property


Public Function SaveToString() As String
    Dim keys() As Variant
    Dim i As Long
    Dim s As String
    keys = mData.keys
    For i = 0 To UBound(keys, 1)
        If Not IsObject(mData(keys(i))) Then
            s = s & keys(i) & ": " & mData(keys(i)) & ", "
        End If
    Next i
    If Len(s) > 2 Then
        SaveToString = Left$(s, Len(s) - 2)
    End If
End Function



Private Sub DoEvent(ByVal Method As String, ParamArray args() As Variant)
    On Error GoTo Err_Init
    Dim SuperClass As clsGeneric
    Dim dict As Dictionary
    Dim EventName As String
    Dim MethodName  As String
    
    ' Make sure this is an instance, not a class
    If mData.Exists("_SuperClass") = False Then
        MsgBox "You can't call a method of the class, only INSTANCES of that class!", vbCritical
        End
    End If
    
    ' Retrieve the superclass methods
    Set SuperClass = mData("_SuperClass")
    Set dict = SuperClass.z_Dictionary
    
    ' Look up this event in the methods
    EventName = "On" & Method
    If dict.Exists(EventName) = False Then
        Form1.Display dict("_ClassName") & " doesn't handle event '" & Method & "'!"
        Exit Sub
    End If
    MethodName = dict(EventName)
    
    ' Execute the specified function
    Select Case UBound(args, 1):
        Case Is < 0:    CallByName Me, MethodName, VbMethod
        Case 0:         CallByName Me, MethodName, VbMethod, args(0)
        Case 1:         CallByName Me, MethodName, VbMethod, args(0), args(1)
        Case 2:         CallByName Me, MethodName, VbMethod, args(0), args(1), args(2)
        Case 3:         CallByName Me, MethodName, VbMethod, args(0), args(1), args(2), args(3)
        Case 4:         CallByName Me, MethodName, VbMethod, args(0), args(1), args(2), args(3), args(4)
        Case 5:         CallByName Me, MethodName, VbMethod, args(0), args(1), args(2), args(3), args(4), args(5)
        Case 6:         CallByName Me, MethodName, VbMethod, args(0), args(1), args(2), args(3), args(4), args(5), args(6)
        Case 7:         CallByName Me, MethodName, VbMethod, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7)
        Case Else:      MsgBox "Increase the number of cases in DoMethod!", vbCritical
                        End
    End Select
            
    Exit Sub
    
Err_Init:
    If Err.Number = 438 Then
        Form1.Display "ERROR - Function " & MethodName & " not found!"
    Else
        MsgBox Err.Number & " - " & Err.Description, vbCritical
        End
    End If
End Sub


'--------------------------------------------------------------------------------
' Events
'--------------------------------------------------------------------------------

Public Sub Speak()
    DoEvent "Speak"
End Sub

Public Sub Gallop(ByVal Speed As Long)
    DoEvent "Gallop", Speed
End Sub

Public Sub Fly(ByVal AirSpeedVelocity As Long)
    DoEvent "Fly", AirSpeedVelocity
End Sub


'--------------------------------------------------------------------------------
' Methods
'--------------------------------------------------------------------------------
Public Sub z_Horse_Gallop(ByVal Speed As Long)
    Display "'" & Me.Name & "' is galloping!  " & Speed & " miles per hour."
End Sub

Public Sub z_Horse_Speak()
    Display "The " & Me.Color & " horse '" & Me.Name & "' whinnies!"
End Sub

Public Sub z_Bird_Speak()
    Display "The " & Me.Color & " bird '" & Me.Name & "' chirps!"
End Sub

Public Sub z_Bird_Fly(ByVal AirSpeedVelocity As Long)
    Display Me.Name & " flies!  " & AirSpeedVelocity & " feet per second."
End Sub


'--------------------------------------------------------------------------------
' Properties
'--------------------------------------------------------------------------------

Public Property Get Name() As String
    Lookup "Name", Name
End Property

Public Property Let Name(ByVal Value As String)
    Change "Name", Value
End Property

Public Property Get Color() As String
    Lookup "Color", Color
End Property

Public Property Let Color(ByVal Value As String)
    Change "Color", Value
End Property

Public Property Get FeatherCount() As Long
    Lookup "FeatherCount", FeatherCount
End Property

Public Property Let FeatherCount(ByVal Value As Long)
    Change "FeatherCount", Value
End Property


