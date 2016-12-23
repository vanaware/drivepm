VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private pSheetAccess As cSheetsV4
Private pActivePrj As Project
Private pTables As cJobject
Private pAssignmentCols As cJobject
Private pResourceCols As cJobject
Private pTaskCols As cJobject
Private pTaskDependencyCols As cJobject
Private pTimePhasedAssignmentCols As cJobject
Private pTimePhasedTaskCols As cJobject
Private pTimePhasedResourceCols As cJobject

Public Property Get ActivePrj() As Project
    Set ActivePrj = pActivePrj
End Property

Public Property Set ActivePrj(MyProject As Project)
    Set pActivePrj = MyProject
End Property

Private Sub Class_Initialize()

    Set pSheetAccess = New cSheetsV4
    pSheetAccess.setAuthName("sheets").setSheetId (getMySheetId())
    
    Set pTables = New cJobject
    pTables.Init(Nothing).addArray
      
    Set pActivePrj = ActiveProject
    
    Set pAssignmentCols = setAssignmentCols()
    Set pResourceCols = setResourceCols()
    Set pTaskCols = setTaskCols()
    Set pTaskDependencyCols = setTaskDependencyCols()
    Set pTimePhasedAssignmentCols = setTimePhasedAssignmentCols()
    Set pTimePhasedTaskCols = setTimePhasedTaskCols()
    Set pTimePhasedResourceCols = setTimePhasedResourceCols()
    
End Sub

Private Function addTable(name As String) As cJobject
    With pTables.add
        .add "range", name
        .add("values").addArray
    End With
    Set addTable = getTable(name)
End Function

Private Function tableExists(table As String) As Boolean
    Dim cJob As New cJobject
    On Error GoTo errorHandler
    Set cJob = pTables.findInArray("range", table).parent
    tableExists = isSomething(cJob)
    Exit Function
errorHandler:
    tableExists = False
End Function

Private Function getTable(table As String) As cJobject
    Set getTable = pTables.findInArray("range", table).parent.child("values")
End Function

Private Function removeTable(table As String) As cJobject
    Set removeTable = pTables.findInArray("range", table).parent.remove
End Function

Private Function getTables() As cJobject
    Set getTables = pTables
End Function

Public Function setTables(cJob As cJobject) As cProject
    Set pTables = cJob
    fixTablesName
    Set setTables = Me
End Function

Public Sub Teardown()
    pTables.Teardown
    Set pSheetAccess = Nothing
    Set pActivePrj = Nothing
    pAssignmentCols.Teardown
    pResourceCols.Teardown
    pTaskCols.Teardown
End Sub

' Sometimes result comes with range incluing A1 notation cells range and we want just sheet name there.
Private Sub fixTablesName()
    Dim job As cJobject, newvalue As String, start As Long
    'Delete everything after ! sign
    For Each job In pTables.children
        newvalue = job.child("range").value
        start = InStr(newvalue, "!")
        If start > 0 Then
            newvalue = Left(newvalue, start - 1)
            job.child("range").value = newvalue
        End If
    Next job
End Sub

Public Sub backupTables()
    'Debug.Print pTables.stringify
    pSheetAccess.pushDatatoSheets pTables
End Sub

Public Sub restoreTables()
    Set pTables = pSheetAccess.pullDatafromSheets
End Sub

Public Sub pullfromTables()
    'TODO
    Debug.Print pTables.stringify
End Sub

Public Sub pushtoTables()
    Dim MyTask As Task
    Dim MyResource As Resource
    Dim MyAssignment As Assignment
    Dim MyTaskDependency As TaskDependency
    Dim result As New cJobject
    
    setProject
    setTask getTaskID(pActivePrj.ProjectSummaryTask)
    For Each MyTask In pActivePrj.Tasks
        If isSomething(MyTask) Then
            '// Push Tasks to Tables
            setTask getTaskID(MyTask)
            setTimePhasedTask getTaskID(MyTask)
            For Each MyTaskDependency In MyTask.TaskDependencies
                If isSomething(MyTaskDependency) And MyTask <> MyTaskDependency.From Then setTaskDependency getTaskID(MyTask), getTaskID(MyTaskDependency.From)
            Next
            For Each MyAssignment In MyTask.Assignments
                If isSomething(MyAssignment) Then
                    Set MyResource = MyAssignment.Resource
                    '// Push Assignments to Tables
                    setAssignment getTaskID(MyTask), getResourceID(MyResource)
                    setTimePhasedAssignment getTaskID(MyTask), getResourceID(MyResource)
                End If
            Next
        End If
    Next
    
    For Each MyResource In pActivePrj.Resources
        '// Push Resources to Tables
        If isSomething(MyResource) Then
            setResource getResourceID(MyResource)
            setTimePhasedResource getResourceID(MyResource)
        End If
    Next
    
End Sub

Private Sub setTask(TaskID As String)
    Const table = "Tasks"
    Dim MyTask As Task
    Dim Col As cJobject
    
    If Not tableExists(table) Then
        With addTable(table).add.addArray
            For Each Col In pTaskCols.children
                .add , Col.Key
            Next
        End With
    End If
    Set MyTask = getTaskbyID(TaskID)
    If isSomething(MyTask) Then
        With getTable(table).add.addArray
            For Each Col In pTaskCols.children
                .add , getTaskField(MyTask, Col.value)
            Next
        End With
    End If
End Sub

Private Sub setResource(ResourceID As String)
    Const table = "Resources"
    Dim MyResource As Resource
    Dim Col As cJobject
    
    If Not tableExists(table) Then
        With addTable(table).add.addArray
            For Each Col In pResourceCols.children
                .add , Col.Key
            Next
        End With
    End If
    Set MyResource = getResourcebyID(ResourceID)
    If isSomething(MyResource) Then
        With getTable(table).add.addArray
            For Each Col In pResourceCols.children
                .add , getResourceField(MyResource, Col.value)
            Next
        End With
    End If
End Sub

Private Sub setAssignment(TaskID As String, ResourceID As String)
    Const table = "Assignments"
    Dim MyAssignment As Assignment
    Dim Col As cJobject
    
    If Not tableExists(table) Then
        With addTable(table).add.addArray
            For Each Col In pAssignmentCols.children
                .add , Col.Key
            Next
        End With
    End If
    Set MyAssignment = getAssignmentbyID(TaskID, ResourceID)
    If isSomething(MyAssignment) Then
        With getTable(table).add.addArray
            For Each Col In pAssignmentCols.children
                .add , getAssignmentField(MyAssignment, Col.value)
            Next
        End With
    End If

End Sub

Private Sub setTimePhasedAssignment(TaskID As String, ResourceID As String)
    Const table = "TimePhasedAssignment"
    Dim MyAssignment As Assignment
    Dim MyTimeScaleValue As TimeScaleValue
    Dim MyTimeScaleValues As TimeScaleValues
    Dim Col As cJobject
    Dim TimeScaleUnit As Long
    Dim DataType As Long

    If Not tableExists(table) Then
        With addTable(table).add.addArray
            For Each Col In pTimePhasedAssignmentCols.children
                .add , Col.Key
            Next
        End With
    End If
    
    Set MyAssignment = getAssignmentbyID(TaskID, ResourceID)
    If isSomething(MyAssignment) Then
        DataType = PjAssignmentTimescaledData.pjAssignmentTimescaledWork
        TimeScaleUnit = PjTimescaleUnit.pjTimescaleMonths
        Set MyTimeScaleValues = MyAssignment.TimeScaleData(pActivePrj.ProjectSummaryTask.start, pActivePrj.ProjectSummaryTask.Finish, DataType, TimeScaleUnit, 1)
        For Each MyTimeScaleValue In MyTimeScaleValues
            With getTable(table).add.addArray
                For Each Col In pTimePhasedAssignmentCols.children
                        Select Case Col.value
                            Case pjTaskText30
                                .add , MyAssignment.Task.GetField(pjTaskText30)
                            Case pjResourceText30
                                .add , MyAssignment.Resource.GetField(pjResourceText30)
                            Case 1
                                .add , Date2Serial(MyTimeScaleValue.StartDate)
                            Case 2
                                .add , Date2Serial(MyTimeScaleValue.EndDate)
                            Case 3
                                .add , MyTimeScaleValue.Index
                            Case 4
                                .add , DataType
                            Case 5
                                .add , TimeScaleUnit
                            Case 6
                                .add , IIf(MyTimeScaleValue.value = "", 0, MyTimeScaleValue.value)
                            Case Else
                                If Col.value < 200000000 And Col.value > 188000000 Then
                                    .add , MyAssignment.Task.GetField(Col.value)
                                ElseIf Col.value < 206000000 And Col.value > 200000000 Then
                                    .add , MyAssignment.Resource.GetField(Col.value)
                                Else
                                    .add , ""
                                End If
                        End Select
                Next
            End With
        Next MyTimeScaleValue
    End If
End Sub

Private Sub setTimePhasedTask(TaskID As String)
    Const table = "TimePhasedTask"
    Dim MyTask As Task
    Dim MyTimeScaleValue As TimeScaleValue
    Dim MyTimeScaleValues As TimeScaleValues
    Dim Col As cJobject
    Dim TimeScaleUnit As Long
    Dim DataType As Long

    If Not tableExists(table) Then
        With addTable(table).add.addArray
            For Each Col In pTimePhasedTaskCols.children
                .add , Col.Key
            Next
        End With
    End If
    
    Set MyTask = getTaskbyID(TaskID)
    If isSomething(MyTask) Then
        DataType = PjTaskTimescaledData.pjTaskTimescaledCost
        TimeScaleUnit = PjTimescaleUnit.pjTimescaleMonths
        Set MyTimeScaleValues = MyTask.TimeScaleData(pActivePrj.ProjectSummaryTask.start, pActivePrj.ProjectSummaryTask.Finish, DataType, TimeScaleUnit, 1)
        
        '#TODO: Falta descontar valores das Tarefas ou Atribuições filhas senão fica acumulado
        
        For Each MyTimeScaleValue In MyTimeScaleValues
            With getTable(table).add.addArray
                For Each Col In pTimePhasedTaskCols.children
                        Select Case Col.value
                            Case pjTaskText30
                                .add , MyTask.GetField(pjTaskText30)
                            Case 1
                                .add , Date2Serial(MyTimeScaleValue.StartDate)
                            Case 2
                                .add , Date2Serial(MyTimeScaleValue.EndDate)
                            Case 3
                                .add , MyTimeScaleValue.Index
                            Case 4
                                .add , DataType
                            Case 5
                                .add , TimeScaleUnit
                            Case 6
                                .add , IIf(MyTimeScaleValue.value = "", 0, MyTimeScaleValue.value)
                            Case Else
                                .add , MyTask.GetField(Col.value)
                        End Select
                Next
            End With
        Next MyTimeScaleValue
    End If
End Sub

Private Sub setTimePhasedResource(ResourceID As String)
    Const table = "TimePhasedResource"
    Dim MyResource As Resource
    Dim MyTimeScaleValue As TimeScaleValue
    Dim MyTimeScaleValues As TimeScaleValues
    Dim Col As cJobject
    Dim TimeScaleUnit As Long
    Dim DataType As Long

    If Not tableExists(table) Then
        With addTable(table).add.addArray
            For Each Col In pTimePhasedResourceCols.children
                .add , Col.Key
            Next
        End With
    End If
    
    Set MyResource = getResourcebyID(ResourceID)
    If isSomething(MyResource) Then
        DataType = PjResourceTimescaledData.pjResourceTimescaledWork
        TimeScaleUnit = PjTimescaleUnit.pjTimescaleMonths
        Set MyTimeScaleValues = MyResource.TimeScaleData(pActivePrj.ProjectSummaryTask.start, pActivePrj.ProjectSummaryTask.Finish, DataType, TimeScaleUnit, 1)
        For Each MyTimeScaleValue In MyTimeScaleValues
            With getTable(table).add.addArray
                For Each Col In pTimePhasedResourceCols.children
                        Select Case Col.value
                            Case pjResourceText30
                                .add , MyResource.GetField(pjResourceText30)
                            Case 1
                                .add , Date2Serial(MyTimeScaleValue.StartDate)
                            Case 2
                                .add , Date2Serial(MyTimeScaleValue.EndDate)
                            Case 3
                                .add , MyTimeScaleValue.Index
                            Case 4
                                .add , DataType
                            Case 5
                                .add , TimeScaleUnit
                            Case 6
                                .add , IIf(MyTimeScaleValue.value = "", 0, MyTimeScaleValue.value)
                            Case Else
                                .add , MyResource.GetField(Col.value)
                        End Select
                Next
            End With
        Next MyTimeScaleValue
    End If
End Sub

Private Sub setTaskDependency(ParentTaskID As String, FromTaskID As String)
    Const table = "TaskDependencies"
    Dim MyTaskDependency As TaskDependency
    Dim Col As cJobject
    
    If Not tableExists(table) Then
        With addTable(table).add.addArray
            For Each Col In pTaskDependencyCols.children
                .add , Col.Key
            Next
        End With
    End If
    Set MyTaskDependency = getTaskDependencybyID(ParentTaskID, FromTaskID)
    If isSomething(MyTaskDependency) Then
        With getTable(table).add.addArray
            For Each Col In pTaskDependencyCols.children
                .add , getTaskDependencyField(MyTaskDependency, Col.value)
            Next
        End With
    End If

End Sub


Private Sub setProject()
    Const table = "Project"
    Dim MyProject As Project
    Set MyProject = pActivePrj
    
    If Not tableExists(table) Then
        With addTable(table).add.addArray
            .add , "Property"
            .add , "Value"
        End With
    End If
    With getTable(table)
        With .add.addArray
            .add , "ID"
            .add , MyProject.id
        End With
        With .add.addArray
            .add , "UniqueID"
            .add , MyProject.UniqueID
        End With
        With .add.addArray
            .add , "FullName"
            .add , MyProject.FullName
        End With
        With .add.addArray
            .add , "Name"
            .add , MyProject.name
        End With
        With .add.addArray
            .add , "Path"
            .add , MyProject.Path
        End With
        With .add.addArray
            .add , "CodeName"
            .add , MyProject.CodeName
        End With
        With .add.addArray
            .add , "ProjectNotes"
            .add , MyProject.ProjectNotes
        End With
        With .add.addArray
            .add , "StatusDate"
            .add , Date2Serial(MyProject.StatusDate)
        End With
        With .add.addArray
            .add , "ProjectStart"
            .add , Date2Serial(MyProject.ProjectStart)
        End With
        With .add.addArray
            .add , "CurrentDate"
            .add , Date2Serial(MyProject.CurrentDate)
        End With
        With .add.addArray
            .add , "CreationDate"
            .add , Date2Serial(MyProject.CreationDate)
        End With
        With .add.addArray
            .add , "ProjectFinish"
            .add , Date2Serial(MyProject.ProjectFinish)
        End With
        With .add.addArray
            .add , "RevisionNumber"
            .add , MyProject.RevisionNumber 'read only
        End With
        With .add.addArray
            .add , "DisplayProjectSummaryTask"
            .add , MyProject.DisplayProjectSummaryTask
        End With
        With .add.addArray
            .add , "YearLabelDisplay"
            .add , MyProject.YearLabelDisplay
        End With
        With .add.addArray
            .add , "MonthLabelDisplay"
            .add , MyProject.MonthLabelDisplay
        End With
        With .add.addArray
            .add , "WeekLabelDisplay"
            .add , MyProject.WeekLabelDisplay
        End With
        With .add.addArray
            .add , "DayLabelDisplay"
            .add , MyProject.DayLabelDisplay
        End With
        With .add.addArray
            .add , "HourLabelDisplay"
            .add , MyProject.HourLabelDisplay
        End With
        With .add.addArray
            .add , "MinuteLabelDisplay"
            .add , MyProject.MinuteLabelDisplay
        End With

        With .add.addArray
            .add , "WBSCodeGenerate"
            .add , MyProject.WBSCodeGenerate
        End With
        With .add.addArray
            .add , "WBSVerifyUniqueness"
            .add , MyProject.WBSVerifyUniqueness
        End With
        With .add.addArray
            .add , "DaysPerMonth"
            .add , MyProject.DaysPerMonth
        End With
        With .add.addArray
            .add , "HoursPerDay"
            .add , MyProject.HoursPerDay
        End With
        With .add.addArray
            .add , "HoursPerWeek"
            .add , MyProject.HoursPerWeek
        End With
        With .add.addArray
            .add , "ScheduleFromStart"
            .add , MyProject.ScheduleFromStart
        End With
    End With
End Sub

Private Function getTaskbyID(TaskID As String) As Task
    Dim MyTask As Task
    For Each MyTask In pActivePrj.Tasks
        If isSomething(MyTask) Then
            If getTaskID(MyTask) = TaskID Then
                Set getTaskbyID = MyTask
                Exit Function
            End If
        End If
    Next
    Set MyTask = pActivePrj.ProjectSummaryTask
    If getTaskID(MyTask) = TaskID Then
        Set getTaskbyID = MyTask
        Exit Function
    End If
    Set getTaskbyID = Nothing
End Function

Private Function getResourcebyID(ResourceID As String) As Resource
    Dim MyResource As Resource
    For Each MyResource In pActivePrj.Resources
        If isSomething(MyResource) Then
            If getResourceID(MyResource) = ResourceID Then
                Set getResourcebyID = MyResource
                Exit Function
            End If
        End If
    Next
    Set getResourcebyID = Nothing
End Function

Private Function getAssignmentbyID(TaskID As String, ResourceID As String) As Assignment
    Dim MyTask As Task
    Dim MyResource As Resource
    Dim MyAssignment As Assignment
    Set MyTask = getTaskbyID(TaskID)
    
    For Each MyAssignment In MyTask.Assignments
        If isSomething(MyAssignment) Then
            Set MyResource = MyAssignment.Resource
            If getResourceID(MyResource) = ResourceID Then
                Set getAssignmentbyID = MyAssignment
                Exit Function
            End If
        End If
    Next
    Set getAssignmentbyID = Nothing
End Function

Private Function getTaskDependencybyID(ParentTaskID As String, FromTaskID As String) As TaskDependency
    Dim MyParent As Task
    Dim MyFrom As Task
    Dim MyTaskDependency As TaskDependency
    
    Set MyParent = getTaskbyID(ParentTaskID)
    
    For Each MyTaskDependency In MyParent.TaskDependencies
        If isSomething(MyTaskDependency) Then
            Set MyFrom = MyTaskDependency.From
            If getTaskID(MyFrom) = FromTaskID Then
                Set getTaskDependencybyID = MyTaskDependency
                Exit Function
            End If
        End If
    Next
    Set getTaskDependencybyID = Nothing
End Function

Private Function getTaskID(MyTask As Task) As String
    Dim TaskID As String
    TaskID = MyTask.GetField(pjTaskText30)
    If TaskID = "" Then
        setTaskID MyTask, genGuid
        TaskID = MyTask.GetField(pjTaskText30)
    End If
    getTaskID = TaskID
End Function

Private Function getResourceID(MyResource As Resource) As String
    Dim ResourceID As String
    ResourceID = MyResource.GetField(pjResourceText30)
    If ResourceID = "" Then
        setResourceID MyResource, genGuid
        ResourceID = MyResource.GetField(pjResourceText30)
    End If
    getResourceID = ResourceID
End Function

Private Sub setTaskID(MyTask As Task, Optional Guid As String = vbNullString)
    If Guid = vbNullString Then
        MyTask.SetField pjTaskText30, genGuid
    Else
        MyTask.SetField pjTaskText30, Guid
    End If
End Sub

Private Sub setResourceID(MyResource As Resource, Optional Guid As String = vbNullString)
    If Guid = vbNullString Then
        MyResource.SetField pjResourceText30, genGuid
    Else
        MyResource.SetField pjResourceText30, Guid
    End If
End Sub

Private Function setTaskCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("UUID (pjTaskText30)", pjTaskText30).add "read-only", True
        .add("pjTaskGuid", pjTaskGuid).add "read-only", True
        .add("pjTaskUniqueID", pjTaskUniqueID).add "read-only", True
        .add("pjTaskID", pjTaskID).add "read-only", True
        .add("pjTaskSummary", pjTaskSummary).add "read-only", True
        .add("pjTaskOutlineNumber", pjTaskOutlineNumber).add "read-only", True
        .add("pjTaskOutlineLevel", pjTaskOutlineLevel).add "read-only", True
        .add("pjTaskParentTask", pjTaskParentTask).add "read-only", True
        .add("pjTaskWBS", pjTaskWBS).add "read-only", True
        .add("pjTaskName", pjTaskName).add "read-only", True
        .add("pjTaskDuration", pjTaskDuration).add "read-only", False           'Read-only for summary tasks
        .add("pjTaskEstimated", pjTaskEstimated).add "read-only", True
        .add("pjTaskUniquePredecessors", pjTaskUniquePredecessors).add "read-only", True
        .add("pjTaskUniqueSuccessors", pjTaskUniqueSuccessors).add "read-only", True
        .add("pjTaskStart", pjTaskStart).add "read-only", True
        .add("pjTaskFinish", pjTaskFinish).add "read-only", True
        .add("pjTaskCreated", pjTaskCreated).add "read-only", True
        .add("pjTaskDeadline", pjTaskDeadline).add "read-only", True
        .add("pjTaskConstraintDate", pjTaskConstraintDate).add "read-only", True
        .add("pjTaskConstraintType", pjTaskConstraintType).add "read-only", True
        .add("pjTaskCalendar", pjTaskCalendar).add "read-only", True
        .add("pjTaskPercentComplete", pjTaskPercentComplete).add "read-only", True
        .add("pjTaskPercentWorkComplete", pjTaskPercentWorkComplete).add "read-only", True
        .add("pjTaskPhysicalPercentComplete", pjTaskPhysicalPercentComplete).add "read-only", True
        .add("pjTaskUpdateNeeded", pjTaskUpdateNeeded).add "read-only", True
        .add("pjTaskFixedCost", pjTaskFixedCost).add "read-only", True
        .add("pjTaskEffortDriven", pjTaskEffortDriven).add "read-only", True
        .add("pjTaskType", pjTaskType).add "read-only", True
        .add("pjTaskFixedDuration", pjTaskFixedDuration).add "read-only", True
        .add("pjTaskMilestone", pjTaskMilestone).add "read-only", True
    End With
    Set setTaskCols = Cols
End Function

Private Function getTaskField(MyTask As Task, FieldID As Long) As Variant
    Select Case FieldID
        Case pjTaskSummary
            getTaskField = MyTask.Summary
        Case pjTaskDuration
            getTaskField = MyTask.Duration 'in minutes
        Case pjTaskStart
            getTaskField = Date2Serial(MyTask.start)
        Case pjTaskFinish
            getTaskField = Date2Serial(MyTask.Finish)
        Case pjTaskCreated
            getTaskField = Date2Serial(MyTask.Created)
        Case pjTaskDeadline
            getTaskField = Date2Serial(MyTask.Deadline)
        Case pjTaskConstraintDate
            getTaskField = Date2Serial(MyTask.ConstraintDate)
        Case pjTaskConstraintDate
            getTaskField = Date2Serial(MyTask.ConstraintDate)
        Case pjTaskFixedCost
            getTaskField = MyTask.FixedCost
        Case pjTaskEstimated
            getTaskField = MyTask.Estimated
        Case pjTaskEffortDriven
            getTaskField = MyTask.EffortDriven
        Case pjTaskFixedDuration
            getTaskField = (MyTask.Type = pjFixedDuration)
        Case pjTaskType
            getTaskField = MyTask.Type
        Case pjTaskParentTask
            getTaskField = getTaskID(MyTask.OutlineParent)
        Case pjTaskUpdateNeeded
            getTaskField = MyTask.UpdateNeeded
        Case pjTaskMilestone
            getTaskField = MyTask.Milestone
        Case pjTaskPercentComplete
            getTaskField = MyTask.PercentComplete
        Case pjTaskPercentWorkComplete
            getTaskField = MyTask.PercentWorkComplete
        Case pjTaskPhysicalPercentComplete
            getTaskField = MyTask.PhysicalPercentComplete
        Case Else
            getTaskField = MyTask.GetField(FieldID)
    End Select
End Function


Private Function setResourceCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("UUID (pjResourceText30)", pjResourceText30).add "read-only", True
        .add("pjResourceGuid", pjResourceGuid).add "read-only", True
        .add("pjResourceUniqueID", pjResourceUniqueID).add "read-only", True
        .add("pjResourceID", pjResourceID).add "read-only", True
        .add("pjResourceName", pjResourceName).add "read-only", True
        .add("pjResourceBaseCalendar", pjResourceBaseCalendar).add "read-only", True
        .add("pjResourceCreated", pjResourceCreated).add "read-only", True
        .add("pjResourceStandardRate", pjResourceStandardRate).add "read-only", True
        
    End With
    Set setResourceCols = Cols
End Function

Private Function getResourceField(MyResource As Resource, FieldID As Long) As Variant
    Select Case FieldID
        Case pjResourceCreated
            getResourceField = Date2Serial(MyResource.Created)
        Case pjResourceStandardRate
            getResourceField = MyResource.StandardRate
        Case Else
            getResourceField = MyResource.GetField(FieldID)
    End Select
End Function


Private Function setAssignmentCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("Task UUID (pjTaskText30)", pjTaskText30).add "read-only", True
        .add("Resource UUID (pjResourceText30)", pjResourceText30).add "read-only", True
        .add("Created", 1).add "read-only", True
        .add("pjTaskName", pjTaskName).add "read-only", True
        .add("pjResourceName", pjResourceName).add "read-only", True
        .add("Units", 2).add "read-only", False
        .add("Work", 3).add "read-only", False
        .add("Start", 4).add "read-only", False
        .add("Finish", 5).add "read-only", False
        .add("WorkContour", 6).add "read-only", False
        .add("ActualWork", 7).add "read-only", False
        .add("RemainingWork", 8).add "read-only", False
        .add("PercentWorkComplete", 9).add "read-only", False
        .add("ActualStart", 10).add "read-only", False
        .add("ActualFinish", 11).add "read-only", False
        .add("Notes", 12).add "read-only", False
        .add("BaselineWork", 13).add "read-only", False
        .add("BaselineStart", 14).add "read-only", False
        .add("BaselineFinish", 15).add "read-only", False
    End With
    Set setAssignmentCols = Cols
End Function

Private Function getAssignmentField(MyAssignment As Assignment, FieldID As Long) As Variant
    Select Case FieldID
        Case pjTaskText30
            getAssignmentField = MyAssignment.Task.GetField(FieldID)
        Case pjResourceText30
            getAssignmentField = MyAssignment.Resource.GetField(FieldID)
        Case 1
            getAssignmentField = Date2Serial(MyAssignment.Created)
        Case 2
            getAssignmentField = MyAssignment.Units
        Case 3
            getAssignmentField = MyAssignment.Work
        Case 4
            getAssignmentField = Date2Serial(MyAssignment.start)
        Case 5
            getAssignmentField = Date2Serial(MyAssignment.Finish)
        Case 6
            getAssignmentField = MyAssignment.WorkContour
        Case 7
            getAssignmentField = MyAssignment.ActualWork
        Case 8
            getAssignmentField = MyAssignment.RemainingWork
        Case 9
            getAssignmentField = MyAssignment.PercentWorkComplete
        Case 10
            getAssignmentField = Date2Serial(MyAssignment.ActualStart)
        Case 11
            getAssignmentField = Date2Serial(MyAssignment.ActualFinish)
        Case 12
            getAssignmentField = MyAssignment.Notes
        Case 13
            getAssignmentField = MyAssignment.BaselineWork
        Case 14
            getAssignmentField = Date2Serial(MyAssignment.BaselineStart)
        Case 15
            getAssignmentField = Date2Serial(MyAssignment.BaselineFinish)
        Case Else
            If FieldID < 200000000 And FieldID > 188000000 Then
                getAssignmentField = MyAssignment.Task.GetField(FieldID)
            ElseIf FieldID < 206000000 And FieldID > 200000000 Then
                getAssignmentField = MyAssignment.Resource.GetField(FieldID)
            Else
                getAssignmentField = ""
            End If
    End Select
End Function

Private Function setTaskDependencyCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("Task UUID (pjTaskText30)", pjTaskText30).add "read-only", True
        .add("From UUID (pjTaskText30)", 1).add "read-only", False
        .add("pjTaskName", pjTaskName).add "read-only", True
        .add("FromTaskName", 2).add "read-only", True
        .add("Type", 3).add "read-only", False
        .add("Lag", 4).add "read-only", False
        .add("LagType", 5).add "read-only", False
    End With
    Set setTaskDependencyCols = Cols
End Function

Private Function getTaskDependencyField(MyTaskDependency As TaskDependency, FieldID As Long) As Variant
    Select Case FieldID
        Case pjTaskText30
            getTaskDependencyField = MyTaskDependency.parent.GetField(FieldID)
        Case 1
            getTaskDependencyField = MyTaskDependency.From.GetField(pjTaskText30)
        Case 2
            getTaskDependencyField = MyTaskDependency.From.GetField(pjTaskName)
        Case 3
            getTaskDependencyField = MyTaskDependency.Type
        Case 4
            getTaskDependencyField = MyTaskDependency.Lag
        Case 5
            getTaskDependencyField = MyTaskDependency.LagType
        Case Else
            getTaskDependencyField = MyTaskDependency.parent.GetField(FieldID)
    End Select
End Function

Private Function setTimePhasedAssignmentCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("Task UUID (pjTaskText30)", pjTaskText30).add "read-only", True
        .add("Resource UUID (pjResourceText30)", pjResourceText30).add "read-only", True
        .add("pjTaskName", pjTaskName).add "read-only", True
        .add("pjResourceName", pjResourceName).add "read-only", True
        .add("StartDate", 1).add "read-only", True
        .add("EndDate", 2).add "read-only", False
        .add("Index", 3).add "read-only", False
        .add("Type", 4).add "read-only", False
        .add("TimeScaleUnit", 5).add "read-only", False
        .add("Value", 6).add "read-only", False
    End With
    Set setTimePhasedAssignmentCols = Cols
End Function

Private Function setTimePhasedTaskCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("Task UUID (pjTaskText30)", pjTaskText30).add "read-only", True
        .add("pjTaskName", pjTaskName).add "read-only", True
        .add("StartDate", 1).add "read-only", True
        .add("EndDate", 2).add "read-only", False
        .add("Index", 3).add "read-only", False
        .add("Type", 4).add "read-only", False
        .add("TimeScaleUnit", 5).add "read-only", False
        .add("Value", 6).add "read-only", False
    End With
    Set setTimePhasedTaskCols = Cols
End Function

Private Function setTimePhasedResourceCols() As cJobject
    Dim Cols As New cJobject
    With Cols.Init(Nothing).addArray
        .add("Resource UUID (pjResourceText30)", pjResourceText30).add "read-only", True
        .add("pjResourceName", pjResourceName).add "read-only", True
        .add("StartDate", 1).add "read-only", True
        .add("EndDate", 2).add "read-only", False
        .add("Index", 3).add "read-only", False
        .add("Type", 4).add "read-only", False
        .add("TimeScaleUnit", 5).add "read-only", False
        .add("Value", 6).add "read-only", False
    End With
    Set setTimePhasedResourceCols = Cols
End Function

Public Sub forceUUID()
    Dim MyTask As Task
    Dim MyResource As Resource
    
    CustomFieldRename pjCustomTaskText30, "UUID"
    CustomFieldRename pjCustomResourceText30, "UUID"
    
    Set MyTask = pActivePrj.ProjectSummaryTask
    If getTaskID(MyTask) = "" Then setTaskID MyTask
    For Each MyTask In pActivePrj.Tasks
        'Debug.Print MyTask.GetField(pjTaskText30)
        If isSomething(MyTask) Then
            If getTaskID(MyTask) = "" Then
                setTaskID MyTask
            End If
        End If
    Next
 
    For Each MyResource In ActiveProject.Resources
        'Debug.Print MyTask.GetField(pjTaskText30)
        If isSomething(MyResource) Then
            If getResourceID(MyResource) = "" Then
                setResourceID MyResource
            End If
        End If
    Next
    
End Sub

Public Property Get self() As cProject
    Set self = Me
End Property
