Sub HeadsUpCreation()
'Open Outlook Application
Dim olApp As Outlook.Application
Set olApp = Outlook.Application

'Create MailObject, the heads up email template will load here
Dim olMsg As Object
'Set the path for the heads up email template .msg file
templatePath = "G:\Opening Team\2_Heads Up\HeadsUpTemplate.msg"
'Open the template file
Set olMsg = olApp.Session.OpenSharedItem(templatePath)
'In order to keep from corrupting the template resave in win temp folder
olMsg.SaveAs "C:\Windows\Temp\HeadsUp.msg"
Set olMsg = olApp.Session.OpenSharedItem("C:\Windows\Temp\HeadsUp.msg")

'The following two commands maybe necessary to fully initialize Outlook automation
' Get a session object.
Dim olNs As Outlook.Namespace
Set olNs = olApp.GetNamespace("MAPI")
' Create an instance of the Inbox folder.
' If Outlook is not already running, this has the side effect of initializing MAPI.
Dim mailFolder As Outlook.Folder
Set mailFolder = olNs.GetDefaultFolder(olFolderInbox)

'Get the input values for the caseName, claimNum, hlNum
caseName = InputBox("Input the Case Name")
claimNum = InputBox("Input the Claim Number")
hlNum = InputBox("Input the HL Number")

'Create Subject line string  Example( Heads Up, CaseName; Claim No. #### ; HL No. #### ; )
subj = "Heads Up, " & caseName & "; Claim No. " & claimNum & "; HL No. " & hlNum
'Create the link to the adjuster file
adjLink = "<a href='\\holden-fs01\Common\Cases\" & caseName & " " & hlNum & "\1.2 Adjuster File'>Adjuster Folder Link</a>"

'Set email values
'[Lookup][D5] = Lead Email, [Lookup][D9] = Senior Email, [Lookup][D13] = Paralegal Email
'[Lookup][D6] = Lead Assistant Email, [Lookup][D10] = Senior Assistant Email, [Lookup][D14] = Paralegal Assistant Email
'Add to CC line, "Joe Cromer <JoeCromer@HoldenLitigation.com>; Cyndi  Russell <CyndiRussell@HoldenLitigation.com>; File Room <FileRoom@holdenlitigation.com>"
With olMsg
    .To = ActiveWorkbook.Sheets("Lookup").Range("D5") & "; " & ActiveWorkbook.Sheets("Lookup").Range("D9") & "; " & ActiveWorkbook.Sheets("Lookup").Range("D13")
    .CC = ActiveWorkbook.Sheets("Lookup").Range("D6") & "; " & ActiveWorkbook.Sheets("Lookup").Range("D10") & "; " & ActiveWorkbook.Sheets("Lookup").Range("D14") & "; Joe Cromer <JoeCromer@HoldenLitigation.com>; Cyndi  Russell <CyndiRussell@HoldenLitigation.com>; File Room <FileRoom@holdenlitigation.com>"
    .Subject = subj
    .Recipients.ResolveAll
    .Display
End With

'Load string to variable for multiple replacements, then exchange placeholders with the employees initial strings
'[Lookup][B5] = Lead Inits
msgBdy = Replace(olMsg.HTMLBody, "Leadinits", UCase(ActiveWorkbook.Sheets("Lookup").Range("B5")))
'[Lookup][B9] = Senior Inits
msgBdy = Replace(msgBdy, "Srinits", UCase(ActiveWorkbook.Sheets("Lookup").Range("B9")))
'[Lookup][B13] = Paralegal Inits
msgBdy = Replace(msgBdy, "Parainits", UCase(ActiveWorkbook.Sheets("Lookup").Range("B13")))
'[Lookup][B14] = Paralegal Assistant Inits
msgBdy = Replace(msgBdy, "ParaAssistinits", UCase(ActiveWorkbook.Sheets("Lookup").Range("B14")))
'adjLink = LinkAdjusterFile
msgBdy = Replace(msgBdy, "LinkAdjusterFile", adjLink)

'Replace actual message body with string
olMsg.HTMLBody = msgBdy

End Sub