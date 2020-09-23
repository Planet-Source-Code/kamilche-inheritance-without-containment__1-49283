Attribute VB_Name = "modMain"
Option Explicit

Public CHorse As clsGeneric
Public CBird As clsGeneric
Public CPegasus As clsGeneric

Public Sub Main()
    Form1.Show
    
    ' Create the classes
    CreateClasses
    
    ' Create and test instances of those objects
    TestObjects
End Sub


Public Sub Display(ByVal s As String)
    Form1.Display s
End Sub


Private Sub CreateClasses()
    ' Horse - default color brown, inherits nothing
    Set CHorse = CreateClass("Horse", _
                "color", "brown", _
                "Name", "HORSE", _
                "OnSpeak", "z_Horse_Speak", _
                "OnGallop", "z_Horse_Gallop")
    
    ' Bird - default color white, inherits nothing
    Set CBird = CreateClass("Bird", _
                "color", "white", _
                "Name", "BIRD", _
                "FeatherCount", 1000, _
                "OnSpeak", "z_Bird_Speak", _
                "OnFly", "z_Bird_Fly")
    
    ' Pegasus - Inherits from both horse and bird.
    '           The last one in 'wins', and overrides the other stats.
    '           Specify unique stats AFTER you're done specifying the inheritance,
    '           else you'll overwrite your custom stats!
    Set CPegasus = CreateClass("Pegasus", _
                "Inherit", CHorse, _
                "Inherit", CBird, _
                "Name", "PEGASUS")

End Sub


Private Function CreateClass(ByVal ClassName As String, ParamArray args() As Variant) As clsGeneric
    ' Create a new class with default properties AND functions
    Dim NewClass As clsGeneric
    Dim i As Long
    Dim Name As String
    
    ' Create the new class
    Set NewClass = New clsGeneric
    
    ' Set the class name
    NewClass.Stat("_ClassName") = ClassName
    For i = 0 To UBound(args, 1)
        If i Mod 2 <> 0 Then
            Name = args(i - 1)
            If StrComp(Name, "Inherit", vbTextCompare) = 0 Then
                NewClass.Inherit args(i)
            ElseIf StrComp(Left$(Name, 1), "_", vbTextCompare) = 0 Then
                MsgBox "Properties can't start with an underscore!", vbCritical
                End
            ElseIf IsObject(args(i)) Then
                Set NewClass.Stat(Name) = args(i)
            Else
                NewClass.Stat(Name) = args(i)
            End If
        End If
    Next i
    
    Display "CLASS " & ClassName & vbCrLf & vbTab & NewClass.SaveToString
    Set CreateClass = NewClass
End Function


Private Function Create(ByRef SuperClass As clsGeneric, ParamArray args() As Variant) As clsGeneric
    ' Create an object (an instance of a class)
    Dim obj As clsGeneric
    Dim i As Long
    Dim Name As String
    
    ' Create a new object
    Set obj = New clsGeneric
    
    ' Set its type
    obj.Stat("_SuperClass") = SuperClass
    
    ' Set initial properties
    For i = 0 To UBound(args, 1)
        If i Mod 2 <> 0 Then
            Name = args(i - 1)
            If StrComp(Name, "Inherit", vbTextCompare) = 0 Then
                MsgBox "You can't specify 'Inherit' in an object instance, only in a class!", vbCritical
                End
            ElseIf StrComp(Left$(Name, 1), "_", vbTextCompare) = 0 Then
                MsgBox "Properties can't start with an underscore!", vbCritical
                End
            ElseIf IsObject(args(i)) Then
                Set obj.Stat(Name) = args(i)
            Else
                obj.Stat(Name) = args(i)
            End If
        End If
    Next i
    Set Create = obj
    
End Function


Private Sub TestObjects()
    Dim Horse As clsGeneric
    Dim Bird As clsGeneric
    Dim Pegasus As clsGeneric
    
    ' Create 3 objects
    Set Horse = Create(CHorse)
    Set Bird = Create(CBird)
    Set Pegasus = Create(CPegasus)
    
    ' Display default attributes
    Display vbCrLf & "Default attributes:"
    Horse.Speak
    Bird.Speak
    Pegasus.Speak
    
    ' Override some attributes
    Horse.Name = "Sarah's Spotted Pony"
    Bird.Name = "Pretty Bird"
    Bird.FeatherCount = 3
    Pegasus.Name = "Pegasus"
    
    ' Display new attributes
    Display vbCrLf & "New names:"
    Horse.Speak
    Bird.Speak
    Pegasus.Speak
    
    ' Custom attributes
    Display vbCrLf & "Custom attributes:"
    Display "The bird has " & Bird.FeatherCount & " feathers."
    Display "The horse has " & Horse.FeatherCount & " feathers."
    Display "The pegasus has " & Pegasus.FeatherCount & " feathers."
    
    ' Call methods that require parameters
    Display vbCrLf & "Call methods with parameters:"
    Horse.Gallop 29
    Horse.Fly -1
    Bird.Gallop 4
    Bird.Fly 35
    Pegasus.Gallop 42
    Pegasus.Fly 5
    
    Horse.Stat("color") = "a different color"
    
End Sub



