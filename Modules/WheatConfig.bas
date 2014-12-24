Attribute VB_Name = "WheatConfig"
'# Wheat Configuration File
'# My answer to persistence configuration of an Excel file
'# As a side note this file should not be exported or imported since this is a local configuration

'# Currently right now, I only need a few options. So I'll stick with it
'# Might expand as an option module

'# PROJECT REPO
'# The name of the project folder, an absolute or relative path.
Public Const PROJECT_REPO As String = "wheat-src"

'# IGNORE_MODULE
'# An array of module names
'# Ignore wheat lib as well
Public IGNORE_MODULE As Variant

Public Sub InitializeVariables()
    IGNORE_MODULE = Array("WheatLib", "Wheat")
End Sub

