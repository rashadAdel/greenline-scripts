input() {
    UserInput := InputBox('courier sheet from excel', 'Please enter a count orders.', , 200)
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
        Send '^c'
        Sleep 300
        Send '^{PgDn}'
        Send '^v'
        Send '{Tab}'
        Send '{Enter}'
        Send '{Escape}'
        Send '+{Tab}'
        Send '^{PgUp}'
        Send '{Down}'
    }
}

^Enter:: {
    courier_sheet()
}