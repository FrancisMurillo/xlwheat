Attribute VB_Name = "ChipList"
Public RepositoryList As Variant

Public Sub ReloadRepositoryList()
    ' This is a list of tuples containing the following
    ' Chip Name, URL
    RepositoryList = Array( _
          Array("Vase", "https://github.com/FrancisMurillo/vase/raw/0.1-poc/vase-RELEASE.xlsm") _
        , Array("Wheat", "https://github.com/FrancisMurillo/wheat/raw/0.1-poc/wheat-RELEASE.xlsm") _
    )
End Sub

'# Gets the designated URL by giving the name of the repo
'# If it does not exist, it returns an empty string
Public Function FindURLByName(ChipName As String) As String
    ReloadRepositoryList
    FindURLByName = ""

    Dim Tuple As Variant
    For Each Tuple In RepositoryList
        If Tuple(0) = ChipName Then FindURLByName = Tuple(1)
    Next
End Function

'# Gets the names of the available chip repos
Public Function ListChipNames() As Variant
    ReloadRepositoryList

    Dim Chips As Variant, Index As Integer
    Chips = Array()
    ReDim Chips(0 To UBound(RepositoryList))
    
    For Index = 0 To UBound(Chips)
        Chips(Index) = RepositoryList(Index)(0)
    Next
    
    ListChipNames = Chips
End Function


