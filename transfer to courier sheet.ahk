input()
{
    InputBox, UserInput, courier sheet from excel, Please enter a count orders.,  , 200, 200
    if ErrorLevel
        MsgBox, CANCEL was pressed.
    else if !RegExMatch(UserInput, "^\d+$"){
        MsgBox, Please enter a valid number.
        input()
    }
    else
        return UserInput

}

courier_sheet(){
    number := input()
    Loop, %number%
    {
        Send, ^c
        Sleep, 300
        send, ^{PgDn}
        send, ^v
        send, {Tab}
        send, {Enter}
        Send, {Escape}
        send, +{Tab}
        send, ^{PgUp}
        Send, {Down}
    }
}
^Enter::
    courier_sheet()
return
tst
