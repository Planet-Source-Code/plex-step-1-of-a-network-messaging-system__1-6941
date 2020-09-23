VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts!"
   ClientHeight    =   5400
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMessages 
      Caption         =   "Messages"
      Height          =   5295
      Left            =   3840
      TabIndex        =   19
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdMsgSave 
         Caption         =   "&Save Message"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txtMessage 
         Height          =   4575
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame fraContacts 
      Caption         =   "Contacts"
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton cmdNewContact 
         Caption         =   "&New"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtICQ 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtlName 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtfName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cmbContact 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "ICQ Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Web Site:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Email Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Fax Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Phone Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Last Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "First Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************
'Application Title  -       Address Book
'Author             -       Thomas Michael
'Purpose            -       Address Book With Reminder Function.
'Created            -       April 1st, 2000
'***************************************************************************************
'Notes:
'This application includes example code to connect to an access database,
'add, modify and delete data from a database, search a database and other
'database funtions. This app gives a working example of the types of apps
'you can create using MS Visual Basic and MS Access. It is fully commented
'so almost any help you need can be found in the comments. I havent really added
'much error checking or handling but hopefully this wont produce any:).
'I hope this example helps you with what you were looking for and if you have
'any questions or comments you can email me at: plexxonic@softhome.net In the next
'Version, I will add the inbox feature so the user can see if they have any new
'Messages.
'***************************************************************************************

Private Sub cmbContact_Click()

    
    'In this Event Sub we will fill the various text
    'fields with the appropriate data from the database
    'When you choose one of the contacts from the combobox
    
    
    'Get all Selected Contacts Information From The database.
    'There is some code in this line that seperates ths first
    'name from the last for we can search the record set by
    'first name.
    Set RS = DB.OpenRecordset("SELECT * FROM tbl_Contacts WHERE tbl_Contacts.fld_fName = " + Chr$(34) + Mid(cmbContact.Text, 1, (InStr(cmbContact.Text, ",") - 1)) + Chr$(34) + ";")
    
    
    
    'Fill In The Various Fields
    With RS
    
        'The First Name Text Box
        txtfName = .Fields("fld_fName")
        'The Last Name Text Box
        txtlName = .Fields("fld_lName")
        'The Phone Number
        txtPhone = .Fields("fld_Phone")
        'The Fax Number
        txtFax = .Fields("fld_Fax")
        'The Email
        txtEmail = .Fields("fld_Email")
        'The Web Site
        txtURL = .Fields("fld_URL")
        'The ICQ Number
        txtICQ = .Fields("fld_ICQ")
        
    End With
    
End Sub

Private Sub cmdDelete_Click()

    'This sub will delete the current contacts information
    'from the screen and the database.
    
    'Select The Current Contacts Information From The Database
    Set RS = DB.OpenRecordset("SELECT * FROM tbl_Contacts WHERE tbl_Contacts.fld_fName = " + Chr$(34) + Mid(cmbContact.Text, 1, (InStr(cmbContact.Text, ",") - 1)) + Chr$(34) + ";")
    
    
    'Delete The Info From The Database.
    With RS
        .Delete
    End With
    
    'Call The ListContacts Sub So We can refresh the contacts
    'That are listed in the combobox.
    Call ListContacts
    
        


End Sub



Private Sub cmdMsgSave_Click()

    'This sub will save the message in the textbox
    'To the current selected user in the database
    
    'Check to make sure there is something in the message and if
    'there is add the message to the database.
    If txtMessage > "" Then
        'There is a message so add it
        
        'Set The recordset object
        Set RS = DB.OpenRecordset("SELECT * FROM tbl_Messages")
        
        'Move To The First Record
        RS.MoveFirst
        
        'Add the message
        With RS
            
            'Let the database know we are adding a new message
            .AddNew
            'The Person it is to
            .Fields("fld_To") = cmbContact.Text
            'Todays date
            .Fields("fld_Date") = Date
            'The Message
            .Fields("fld_Message") = Trim(txtMessage)
            
            'Update the recordset
            .Update
        
        End With
    
        'Let the user know it was added successfully
        MsgBox "Message Logged Successfully!!"
        
        'Clear the Text Box
        txtMessage = ""
        
        Else
            'Dont do anything since there is no message
        End If

End Sub

Private Sub cmdNewContact_Click()

    'This sub will gather information from the user to create
    'a new contact and then refesh the contacts list. We
    'will gather info from the user via input boxes.
    
    'Open The Record Set So We Can add a new contact.
    Set RS = DB.OpenRecordset("SELECT * FROM tbl_Contacts")
    
    'Move to the last record
    RS.MoveLast
    
    'Gather new contact information from the user and add it to
    'the database and the update it.
    
    With RS
    
        'Let the record set know we are going to add a new record
        .AddNew
    
        'The first name
        .Fields("fld_fName") = Trim(InputBox("Contacts First Name", "New Contact"))
        'The Last name
        .Fields("fld_lName") = Trim(InputBox("Contacts Last Name", "New Contact"))
        'The Phone number
        .Fields("fld_Phone") = Trim(InputBox("Contacts Phone Number", "New Contact"))
        'The Fax Number
        .Fields("fld_Fax") = Trim(InputBox("Contacts Fax Number", "New Contact"))
        'The Email address
        .Fields("fld_Email") = Trim(InputBox("Contacts Email Address", "New Contact"))
        'The Web URL
        .Fields("fld_URL") = Trim(InputBox("Contacts URL", "New Contact"))
        'The ICQ Number
        .Fields("fld_ICQ") = Trim(InputBox("Contacts ICQ Number", "New Contact"))
        
        'Update The Record Set
        .Update
        
    End With
    
    'Let the user know the new contact was added.
    MsgBox "New Contact Added!!"
    
    'Update The Combobox of Contacts.
    Call ListContacts
        
        
        


End Sub

Private Sub cmdUpdate_Click()

    'This sub will update the information in the database with any
    'information the user has changed in the text boxes.
    
    'Get all Selected Contacts Information From The database.
    'There is some code in this line that seperates ths first
    'name from the last for we can search the record set by
    'first name.
    Set RS = DB.OpenRecordset("SELECT * FROM tbl_Contacts WHERE tbl_Contacts.fld_fName = " + Chr$(34) + Mid(cmbContact.Text, 1, (InStr(cmbContact.Text, ",") - 1)) + Chr$(34) + ";")
    
    'Change the information and then update the record
    With RS
        
        .Edit
        'The First Name
        .Fields("fld_fName") = Trim(txtfName)
        'The Last Name
        .Fields("fld_lName") = Trim(txtlName)
        'The Phone Number
        .Fields("fld_Phone") = Trim(txtPhone)
        'The Fax Number
        .Fields("fld_Fax") = Trim(txtFax)
        'The Email Address
        .Fields("fld_Email") = Trim(txtEmail)
        'The Website URL
        .Fields("fld_URL") = Trim(txtURL)
        'The ICQ Number
        .Fields("fld_ICQ") = Trim(txtICQ)
        
        'Now Update The Modified Record in the Database. Whenever you
        'Modify info in a record you must update it so the changes will
        'Take place.
        .Update
        
        'Let the user know the update was successful
        MsgBox "Update Successful!!"
        
        'Refresh the combobox
        Call ListContacts
        
    End With

End Sub

Private Sub Form_Load()
    
    'Call The Sub That Connects Us To The Database.
    Call dbConnect(App.Path & "\data\", "data.mdb")
    
    'Call The sub that will add all of the Contacts
    'To The Contacts Combo Box.
    Call ListContacts
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'This sub event will make sure we have cleared all
    'of our refrences and set them to nothing to clear up
    'Memory.
    
    'Close The Recordset
    RS.Close
    'Close the connection to the database
    DB.Close
    
    'Clear memory used by the form
    Set frmMain = Nothing
    'Clear memory used by the record set
    Set RS = Nothing
    'Clear memory used by the database connection.
    Set DB = Nothing

    'Unload The main form
    Unload Me
        
End Sub
