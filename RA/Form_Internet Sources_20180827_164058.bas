VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Internet Sources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub chk_Auction_BeforeUpdate(Cancel As Integer)
    Dim ClientID As Integer
    ClientID = Nz(DLookup("ClientID", "CaseList", "Filenumber = " & Me!FileNumber))
    If (Nz(DLookup("AuctionDotComAllowed", "ClientList", "ClientID =" & ClientID), -1)) >= 0 Then
        MsgBox "This client is not approved for Auction.com, please see your manager.  This setting will not be saved."
        Cancel = True
        Me!chk_Auction.Undo
        
        Dim SendTo As String
        Dim OutlookObj As Object, NewMsg As Object
        SendTo = Nz(DLookup("sValue", "DB", "ID=50"))
        'Nz(DLookup("sValue", "DB", "ID=" & distriputionList))
        'AttachmentPath = DocLocation & DocBucket(FileID) & "\" & FileID & "\Title Order " & Format$(Now(), "yyyymmdd") & " 7.pdf"

        Set OutlookObj = CreateObject("Outlook.Application")
        Set NewMsg = OutlookObj.CreateItem(0)

        'Create new email


        With NewMsg ' this is this test
            .To = SendTo
            '.CC = "tr@acertitle.com"
            .Body = "File: " & Me!FileNumber & " is flagged in client system, but " & Nz(DLookup("LongClientName", "ClientList", "ClientID =" & ClientID)) & " has not authorized us to release information to Auction.com.  Changes have been rolled back, but please review file with client."
            .Subject = "Auction.com Exception - " & Nz(DLookup("LongClientName", "ClientList", "ClientID =" & ClientID), -1)
            .Send
        End With
    End If
    
Exit_chk_Auction_BeforeUpdate:
        
        
    Exit Sub

Err_chk_Auction_BeforeUpdate:
    MsgBox Err.Description
    Cancel = True
    Me!chk_Auction.Undo
    Resume Exit_chk_Auction_BeforeUpdate
End Sub

Private Sub chk_xHome_BeforeUpdate(Cancel As Integer)
   Dim ClientID As Integer
    ClientID = Nz(DLookup("ClientID", "CaseList", "Filenumber = " & Me!FileNumber))
    If (Nz(DLookup("xHomeAllowed", "ClientList", "ClientID =" & ClientID), -1)) >= 0 Then
        MsgBox "This client is not approved for xHome, please see your manager.  This setting will not be saved."
        Cancel = True
        Me!chk_xHome.Undo
        
        Dim SendTo As String
        Dim OutlookObj As Object, NewMsg As Object
        SendTo = Nz(DLookup("sValue", "DB", "ID=50"))
 
        Set OutlookObj = CreateObject("Outlook.Application")
        Set NewMsg = OutlookObj.CreateItem(0)

        'Create new email


        With NewMsg ' this is this test
            .To = SendTo
            '.CC = "tr@acertitle.com"
            .Body = "File: " & Me!FileNumber & " is flagged in client system, but " & Nz(DLookup("LongClientName", "ClientList", "ClientID =" & ClientID)) & " has not authorized us to release information to xHome.  Changes have been rolled back, but please review file with client."
            .Subject = "xHome Exception - " & Nz(DLookup("LongClientName", "ClientList", "ClientID =" & ClientID), -1)
            .Send
        End With
    End If
    
Exit_chk_xHome_BeforeUpdate:
        
        
    Exit Sub

Err_chk_xHome_BeforeUpdate:
    MsgBox Err.Description
    Cancel = True
    Me!chk_xHome.Undo
    Resume Exit_chk_xHome_BeforeUpdate
End Sub

Private Sub cmdClose_Click()
DoCmd.Close
End Sub
