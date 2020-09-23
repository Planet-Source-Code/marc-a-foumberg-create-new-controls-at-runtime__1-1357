<div align="center">

## Create New Controls At Runtime \!


</div>

### Description

This code will allow you to create a new instance of a control

at runtime !

Imagine. When your app is RUNNING you can create a new command button

or a textbox on the fly ! The code they didn't want you to know ;)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marc A\. Foumberg](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marc-a-foumberg.md)
**Level**          |Unknown
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marc-a-foumberg-create-new-controls-at-runtime__1-1357/archive/master.zip)

### API Declarations

NONE


### Source Code

```
'Create a new project.
'Add a command button.
'Name the button...
' Command1(0)
'As if it were an aray.
'Its sometimes easyier to create
'an aray to begin with. If you do
'be sure to delete all button except
'Command1(0).
'The Code...
Private Sub Command1_Click(Index As Integer)
Static I As Integer
I = I + 1
Load Command1(I)
Command1(I).Left = Command1(I - 1).Left + 200
Command1(I).Top = Command1(I - 1).Top + 600
Command1(I).Caption = "New Button !"
Command1(I).Visible = True
End Sub
'At runtime this will create a new
'command button.
'To add additional function you could add
'an IF statement. As follows...
Private Sub Command1_Click(Index As Integer)
On Error GoTo Handler1
'Create new button
Static I As Integer
I = I + 1
Load Command(I)
Command(I).Left = 2460
Command(I).Top = 5520
Command(I).Caption = "For Real This Time" ' change the caption
Command(I).Visible = True
' Code to unload the form when the new button is clicked
If Command(1) Then
Unload Me
End If
Handler1:
End Sub
'Email Marc at 3dtech@acwn.com with any questions.
'try this with other controls !
```

