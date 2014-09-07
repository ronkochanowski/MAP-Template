VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    Dim ccExt As ContentControl
    Dim ccDir As ContentControl
    Dim ccMob As ContentControl
    Dim ccEmail As ContentControl
    
    With ActiveDocument
        
    
        Select Case ContentControl.Tag
            Case "Consultant"
                Select Case ContentControl.Range.Text
                    Case "Lesley Baker"
                        .SelectContentControlsByTag("idExtension").Item(1).Range.Text = "8310"
                        .SelectContentControlsByTag("idDirectNumber").Item(1).Range.Text = "843.413.8310"
                        .SelectContentControlsByTag("idMobileNumber").Item(1).Range.Text = "210.274.3441"
                        .SelectContentControlsByTag("idEMail").Item(1).Range.Text = "lesley.baker@acstechnologies.com"
                    Case "David Childers"
                        .SelectContentControlsByTag("idExtension").Item(1).Range.Text = "8206"
                        .SelectContentControlsByTag("idDirectNumber").Item(1).Range.Text = "843.413.8206"
                        .SelectContentControlsByTag("idMobileNumber").Item(1).Range.Text = "770.616.4444"
                        .SelectContentControlsByTag("idEMail").Item(1).Range.Text = "david.childers@acstechnologies.com"
                    Case "Russ Fortier"
                        .SelectContentControlsByTag("idExtension").Item(1).Range.Text = "8193"
                        .SelectContentControlsByTag("idDirectNumber").Item(1).Range.Text = "843.413.8193"
                        .SelectContentControlsByTag("idMobileNumber").Item(1).Range.Text = "843.496.1924"
                        .SelectContentControlsByTag("idEMail").Item(1).Range.Text = "russ.fortier@acstechnologies.com"
                    Case "Jennifer Hollingshead"
                        .SelectContentControlsByTag("idExtension").Item(1).Range.Text = "8205"
                        .SelectContentControlsByTag("idDirectNumber").Item(1).Range.Text = "843.413.8205"
                        .SelectContentControlsByTag("idMobileNumber").Item(1).Range.Text = "210.274.3441"
                        .SelectContentControlsByTag("idEMail").Item(1).Range.Text = "jennifer.Hollingshead@acstechnologies.com"
                    Case "Keith Hudgins"
                        .SelectContentControlsByTag("idExtension").Item(1).Range.Text = "8428"
                        .SelectContentControlsByTag("idDirectNumber").Item(1).Range.Text = "843.413.8428"
                        .SelectContentControlsByTag("idMobileNumber").Item(1).Range.Text = "512.431.2890"
                        .SelectContentControlsByTag("idEMail").Item(1).Range.Text = "keith.hudgins@acstechnologies.com"
                    Case "Ron Kochanowski"
                        .SelectContentControlsByTag("idExtension").Item(1).Range.Text = "8434"
                        .SelectContentControlsByTag("idDirectNumber").Item(1).Range.Text = "843.413.8434"
                        .SelectContentControlsByTag("idMobileNumber").Item(1).Range.Text = "413.896.9360"
                        .SelectContentControlsByTag("idEMail").Item(1).Range.Text = "ronnie.kochanowski@acstechnologies.com"
                    Case "Mark Seymour"
                        .SelectContentControlsByTag("idExtension").Item(1).Range.Text = "8077"
                        .SelectContentControlsByTag("idDirectNumber").Item(1).Range.Text = "843.413.8077"
                        .SelectContentControlsByTag("idMobileNumber").Item(1).Range.Text = "803.856.1043"
                        .SelectContentControlsByTag("idEMail").Item(1).Range.Text = "mark.seymour@acstechnologies.com"
                    Case "Click here to select a Consultant"
                        .SelectContentControlsByTag("idExtension").Item(1).Range.Text = ""
                        .SelectContentControlsByTag("idDirectNumber").Item(1).Range.Text = ""
                        .SelectContentControlsByTag("idMobileNumber").Item(1).Range.Text = ""
                        .SelectContentControlsByTag("idEMail").Item(1).Range.Text = ""
                    Case Else
                End Select
            Case Else
         'The user exited some other content control that we don't care about.
        End Select
    End With
End Sub

    
