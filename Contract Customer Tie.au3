Global $Runner
HotKeySet("{HOME}", "start")
HotKeySet("{END}", "stop")

While 1
    Sleep(1000)
WEnd

Test()

Func start()
$Runner = Not $Runner
While $Runner
   Sleep(300)
WinActivate("Microsoft Excel - Promo")
Sleep(1000)
Send("{DOWN}")
Send("{RIGHT}")
Send("{RIGHT}")
Sleep(1000)
Send("^c") ;Copy the billto customer ID with leading period
Sleep(1000)
WinActivate("Customer Maintenance")
Sleep(1000)
Send("{UP 2}") ;move cursor to customer entry
Sleep(1000)
Send("^v") ;Paste billto Number
Sleep(1000)
Send("{ENTER}") ;Enters search for customer billto
Sleep(100)
WinActivate("Microsoft Excel - Promo")
Sleep(1000)
Send("{LEFT 2}") ;move cursor to contract ID

Sleep(1000)
Send("^c") ;copies contract id
Sleep(1000)
WinActivate("Customer Maintenance")
Sleep(1000)
Send("^{F1}") ;enter contract customer screen
Sleep(2000)
Send("{LEFT}") ;move cursor to the left hand column
Sleep(100)
Send("^{END}") ;Move cursor to bottom of Entry screen
Sleep(1000)

Send("^v") ;paste contract id
Sleep(1000)
Send("{ENTER}")
Sleep(1000)
Send("^s") ;save changes to contract screen
Sleep(1000)
Send("{F12}") ;exit contract screen

WEnd

EndFunc

func stop()
   Exit
   EndFunc
