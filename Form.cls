VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ApplicantEmail As String
Public Applicant As String
Public Admin As String
Public AdminEmail As String
Public HPC As String

Public Property Get Self() As Form
    Set Self = Me
End Property
