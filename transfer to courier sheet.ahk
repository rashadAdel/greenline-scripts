input() {
    UserInput := InputBox('courier sheet from excel', 'Please enter a count orders.', , 1)
    if UserInput.Result = 'Cancel' {
        MsgBox 'CANCEL was pressed.'
    } else if !RegExMatch(UserInput.Value, '^\d+$') {
        MsgBox 'Please enter a valid number.'
        input()
    } else {
        return UserInput.Value
    }
}

courier_sheet() {
    number := input()
    Loop number {
        SendEvent '^c'
        Sleep 300
        SendEvent '^{PgDn}'
        SendEvent '^v'
        SendEvent '{Tab}'
        SendEvent '{Enter}'
        SendEvent '{Escape}'
        SendEvent '+{Tab}'
        SendEvent '{Enter}'
        SendEvent '^{PgUp}'
        SendEvent '{Down}'
    }
}

^Enter:: {
    courier_sheet()
}