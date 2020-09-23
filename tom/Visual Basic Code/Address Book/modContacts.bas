Attribute VB_Name = "modContacts"
'This Lines Sets It So All Variables
'Must Be Declared. Such As (Dim X As Integer)
Option Explicit

    'This is the Database Object Refrence
    Public DB As Database
    'This is the Recordset Object Refrence
    Public RS As Recordset


Public Sub dbConnect(dbPath As String, dbName As String)
    
    'This Sub Will Connect To The Database Specified
    'In The Variables That Were Passed To This Sub
    
    'This Little If... Else Statement Will Make
    'sure that there Is a "\" Backslash at the
    'end of the dbPath String. If Not It will
    'add one on to it.
    
    If Right(dbPath, 1) = "\" Then
        'Do Nothing Becasue The Backslash
        'Is Alredy There
    Else
        'This Add The Backslash To The dbPath
        'String So We dont recieve an error.
        dbPath = dbPath & "\"
    End If
        
    
    'This Line Actually Opens The Database For
    'Useage Later On.
    Set DB = OpenDatabase(dbPath & dbName)

End Sub


Public Sub ListContacts()

    'Clear any existing names form the combobox
    frmMain.cmbContact.Clear

    'This Sub will add every contacts name from the database
    'to the list of contacts.
    
    'This Opens up the record set object and select all
    'the contacts first and last names.
    Set RS = DB.OpenRecordset("SELECT fld_fName,fld_lName FROM tbl_Contacts")
    
    'Move To The First Record
    RS.MoveFirst
    
    'This Joins The First and Last Names Together Then
    'Adds Them To The Contact Combo Box On The Form
    
    With RS
        
        'Loop Through The Records Untill We Reach The Last Record.
        Do While Not RS.EOF
        
            'Add The Items
            frmMain.cmbContact.AddItem .Fields("fld_fName") & ", " & _
                                       .Fields("fld_lName")
        
            'Move To The Next Contact Record
            .MoveNext

        Loop
        
        'Set The Initial Contact Selected To The First Person
        'In The List
        frmMain.cmbContact.ListIndex = 0
    
    End With
    
End Sub
