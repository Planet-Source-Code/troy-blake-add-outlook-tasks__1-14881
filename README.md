<div align="center">

## Add Outlook Tasks


</div>

### Description

Used to show and add new tasks to Outlook. Simple code to demo getting task list and also adding a new task in Outlook. Please vote.
 
### More Info
 
Create Reference to Outlook. Create a form, insert code. Also put a combobox on the form and name it cboTasklist. Load form at runtime to view current task list. Simple code to demo getting task list and also adding a new task. Not very useful as it is listed here, but you can see how you could change it to be a useful function. Please vote.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Troy Blake](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/troy-blake.md)
**Level**          |Beginner
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/troy-blake-add-outlook-tasks__1-14881/archive/master.zip)





### Source Code

```
'Simple Outlook Task View/Add Code
'Troy Blake - Logan's Roadhouse, Inc.
Private Sub InitForm()
 'Loads current task to dropdown, then adds
 'a task for John Smith. John gets the task
 'sent to him via Outlook.
 Dim oApp as Outlook.Application
 Dim oNspc as NameSpace
 Dim oItm as TaskItem
 Dim myItem as TaskItem
 Set oApp = CreateObject("Outlook.Application")
 Set oNspc = oApp.GetNamespace("MAPI")
 For Each oItm in oNspc.GetDefaultFolder(olFolderTasks).Items
  'Loop through all tasks and show subject
  'in dropdown.
  With Me.cboTasklist
   .AddItem (oItm.Subject)
  End With
 Next oItm
 oNspc.GetDefaultFolder(olFolderTasks).Items.Add
 Set myItem = oApp.CreateItem(olTaskItem)
 'Create a new task
 With myItem
  .Subject = "Subject"
  .Assign = "Assign"
  .Body = "Task Body"
  .PercentComplete = 10
  'Set due date for tomorrow
  .DueDate = DateAdd("d",1,Date)
  .ReminderSet = True
  .ReminderTime = "9:00 AM"
  'Outlook name of person to get task
  .Recipients.Add "John Smith"
  .Close (olSave)
 End With
 'Send the task (like email)
 myItem.Send
 Set myItem = Nothing
 Set oItm = Nothing
 Set oNspc = Nothing
 Set oApp = Nothing
End Sub
Private Sub Form_Load()
 'Call out sample sub at form load
 InitForm
End Sub
```

