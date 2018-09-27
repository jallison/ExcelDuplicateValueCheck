Attribute VB_Name = "Performance"
Public screenUpdateState As Boolean
Public statusBarState As Boolean
Public eventsState As Boolean

Function TurnOffAll()

    screenUpdateState = Application.ScreenUpdating
    eventsState = Application.EnableEvents
    statusBarState = Application.DisplayStatusBar
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False

End Function

Function TurnOnAll()

    Application.ScreenUpdating = screenUpdateState
    Application.EnableEvents = eventsState
    Application.DisplayStatusBar = statusBarState

End Function
