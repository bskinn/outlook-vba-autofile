Attribute VB_Name = "AutoFile"
Option Explicit

' =============  B&P  ==============

Sub launchBnPFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "B&P"
    UFAutoFile.baseFldPath = "zArchive\Faraday\B&P"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Proposal Folder"
    UFAutoFile.popDepth = 0
    UFAutoFile.LBxFontSize = 9#
    
    UFAutoFile.Show
End Sub

Sub launchBnPFileOnInspector()
    Load UFAutoFile
    
    Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    
    launchBnPFile
End Sub


' ===========  CLIENTS  ===========

Sub launchClientFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Clients"
    UFAutoFile.baseFldPath = "zArchive\Clients"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Client"
    UFAutoFile.popDepth = 0
    
    UFAutoFile.Show
End Sub

Sub launchClientFileOnInspector()
    Load UFAutoFile
    
    Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    
    launchClientFile
End Sub


' ===========  COLLABORATORS  ===========

Sub launchCollabFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Collaborators"
    UFAutoFile.baseFldPath = "zArchive\Collaborators"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Collaborator"
    UFAutoFile.popDepth = 0
    
    UFAutoFile.Show
End Sub

Sub launchCollabFileOnInspector()
    Load UFAutoFile
    
    Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    
    launchCollabFile
End Sub


' ===========  CONTRACTORS  ===========

Sub launchContractorFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Contractors"
    UFAutoFile.baseFldPath = "zArchive\Contractors"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Contractor"
    UFAutoFile.popDepth = 1
    
    UFAutoFile.Show
End Sub

Sub launchContractorFileOnInspector()
    Load UFAutoFile
    
    Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    
    launchContractorFile
End Sub


' ===========  PROJECTS  ===========

Sub launchProjectFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "~~^\S+"  ' Leading non-whitespace regex
    UFAutoFile.baseFldPath = "zArchive\Faraday\Projects"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Project"
    UFAutoFile.popDepth = 1
    
    UFAutoFile.Show
End Sub

Sub launchProjectFileOnInspector()
    Load UFAutoFile
    
    Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    
    launchProjectFile
End Sub


' ===========  SOCIETIES  ===========

Sub launchSocietyFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Societies"
    UFAutoFile.baseFldPath = "zArchive\Societies"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Society"
    UFAutoFile.popDepth = 0
    
    UFAutoFile.Show
End Sub

Sub launchSocietyFileOnInspector()
    Load UFAutoFile
    
    Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    
    launchSocietyFile
End Sub


' ===========  VENDORS  ===========

Sub launchVendorFile()
    Load UFAutoFile
    
    UFAutoFile.catAddStr = "Vendors"
    UFAutoFile.baseFldPath = "zArchive\Vendors"
    UFAutoFile.storeName = "Outlook"
    UFAutoFile.Caption = "Select Vendor"
    UFAutoFile.popDepth = 1
    
    UFAutoFile.Show
End Sub

Sub launchVendorFileOnInspector()
    Load UFAutoFile
    
    Set UFAutoFile.tgtObj = ActiveInspector.CurrentItem
    
    launchVendorFile
End Sub


