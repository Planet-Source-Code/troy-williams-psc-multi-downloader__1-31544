VERSION 5.00
Begin VB.Form frmInfo 
   Caption         =   "Information"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox txtInfo 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim sInfo As String

sInfo = sInfo & "Written by: Troy Williams" & vbNewLine
sInfo = sInfo & "Email: fenris@hotmail.com" & vbNewLine
sInfo = sInfo & vbNewLine
sInfo = sInfo & vbNewLine
sInfo = sInfo & "First off, the most of the code is original. The rest of the code has been" & vbNewLine
sInfo = sInfo & "put to gether from various sources around the internet including PSC. So if you recognize code that you wrote" & vbNewLine
sInfo = sInfo & "I will be more then happy to put your name to the code." & vbNewLine
sInfo = sInfo & "" & vbNewLine
sInfo = sInfo & "" & vbNewLine
sInfo = sInfo & "" & vbNewLine
sInfo = sInfo & "" & vbNewLine
sInfo = sInfo & "This program is designed to download source code from PSC.." & vbNewLine
sInfo = sInfo & "- It is capable of downloading up to six files at the same time(well actually, the number is virtually unlimited, but more then 4 or 5 will piss PSC off. I found that 2 gives good results)." & vbNewLine
sInfo = sInfo & "- It supports cut and pasting of urls, as well as drag and dropping urls." & vbNewLine
sInfo = sInfo & "I wrote this program because my home computer (win XP pro) could not access the PSC site for some reason." & vbNewLine
sInfo = sInfo & "I was receiving the code of the day newsletter, in which were the links to that days uploads." & vbNewLine
sInfo = sInfo & "So I put two and two together and wrote an app to download the files." & vbNewLine
sInfo = sInfo & "Along the way I discovered a number of interesting things. The urls in the newsletters are redirect links" & vbNewLine
sInfo = sInfo & "So I developed a way to follow the redirect links. The program has a number of interesting features as well." & vbNewLine
sInfo = sInfo & "The urls to download are stored in a listbox, which basically acts like a que. This allows multiple winsock controls to download different files" & vbNewLine
sInfo = sInfo & "When a valid download url is found (i.e. the normal sourcecode download page) it copies the html to a directory named for the <title> of the webpage." & vbNewLine
sInfo = sInfo & "This allows for a description of the code to be saved along with the zip file." & vbNewLine
sInfo = sInfo & "Also the html and file headers are saved in that directory as well as a *.url file with the download url." & vbNewLine
sInfo = sInfo & "" & vbNewLine
sInfo = sInfo & "" & vbNewLine
sInfo = sInfo & "The program also logs urls that caused problems to a log window, which allows the log to be written to disk." & vbNewLine
sInfo = sInfo & "The Url list can be saved to a file as well." & vbNewLine
sInfo = sInfo & "" & vbNewLine
sInfo = sInfo & "" & vbNewLine
sInfo = sInfo & "To use the program...." & vbNewLine
sInfo = sInfo & "Start up the program, either from the ide or from a full compile." & vbNewLine
sInfo = sInfo & "Select options from the menu and locate the directory where you want to download the files to, also enter the number of simultaneous downloads you want to have. " & vbNewLine
sInfo = sInfo & "Goto File/New Downloads" & vbNewLine
sInfo = sInfo & "An URL List window will popup, simply copy and past (or drag and drop) the urls into the listbox" & vbNewLine
sInfo = sInfo & "When you are ready to begin downloading, click on the download button" & vbNewLine

txtInfo = sInfo
End Sub

Private Sub Form_Resize()
txtInfo.Top = Me.ScaleTop
txtInfo.Left = Me.ScaleLeft
txtInfo.Height = Me.ScaleHeight
txtInfo.Width = Me.ScaleWidth
End Sub
