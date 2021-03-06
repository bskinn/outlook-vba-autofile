Attribute VB_Name = "AutoFile"
Option Explicit

' =============  Archived B&P  ==============

Sub launchBnPFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "B&P"
    UFAutoFile.baseFldPath = "zArchive\Faraday\B&P"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Proposal Folder"
    UFAutoFile.popDepth = 0
    UFAutoFile.LBxFontSize = 9#
    
    If TypeOf ActiveWindow Is Inspector Then
        Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    End If
    
    UFAutoFile.Show
End Sub


' =============  Current B&P  ==============

Sub launchActiveBnPFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "B&P"
    UFAutoFile.baseFldPath = "Inbox\4. Active B&P"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Active Proposal Folder"
    UFAutoFile.popDepth = 0
    UFAutoFile.LBxFontSize = 9#
    
    If TypeOf ActiveWindow Is Inspector Then
        Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    End If
    
    UFAutoFile.Show
End Sub




' ===========  CLIENTS  ===========

Sub launchClientFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Clients"
    UFAutoFile.baseFldPath = "zArchive\Clients"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Client"
    UFAutoFile.popDepth = 0
    
    If TypeOf ActiveWindow Is Inspector Then
        Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    End If
    
    UFAutoFile.Show
End Sub


' ===========  COLLABORATORS  ===========

Sub launchCollabFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Collaborators"
    UFAutoFile.baseFldPath = "zArchive\Collaborators"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Collaborator"
    UFAutoFile.popDepth = 1
    
    If TypeOf ActiveWindow Is Inspector Then
        Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    End If
    
    UFAutoFile.Show
End Sub


' ===========  CONTRACTORS  ===========

Sub launchContractorFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Contractors"
    UFAutoFile.baseFldPath = "zArchive\Contractors"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Contractor"
    UFAutoFile.popDepth = 1
    
    If TypeOf ActiveWindow Is Inspector Then
        Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    End If
    
    UFAutoFile.Show
End Sub


' ===========  PROJECTS  ===========

Sub launchProjectFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "~~^\S+"  ' Leading non-whitespace regex
    UFAutoFile.baseFldPath = "zArchive\Faraday\Projects"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Project"
    UFAutoFile.popDepth = 1
    
    If TypeOf ActiveWindow Is Inspector Then
        Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    End If
    
    UFAutoFile.Show
End Sub


' ===========  SOCIETIES  ===========

Sub launchSocietyFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Societies"
    UFAutoFile.baseFldPath = "zArchive\Societies"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Society"
    UFAutoFile.popDepth = 0
    
    If TypeOf ActiveWindow Is Inspector Then
        Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    End If
    
    UFAutoFile.Show
End Sub


' ===========  VENDORS  ===========

Sub launchVendorFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Vendors"
    UFAutoFile.baseFldPath = "zArchive\Vendors"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Vendor"
    UFAutoFile.popDepth = 1
    
    If TypeOf ActiveWindow Is Inspector Then
        Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    End If
    
    UFAutoFile.Show
End Sub


