Attribute VB_Name = "ChipInfo"
Public References As Variant
Public Modules As Variant

Public Sub Initialize()
    References = Array( _
        "Microsoft Visual Basic for Applications Extensibility *", _
        "Microsoft Scripting Runtime")
    Modules = Array( _
        "Wheat", "WheatLib", "WheatConfig", "WheatUtil")
End Sub
