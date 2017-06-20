from pywinauto.application import Application

app = Application().Connect(title=u'Sierra \xb7 Miami University Libraries \xb7 Craig Boman', class_name='Chrome_WidgetWin_1')
chromewidgetwin = app.Chrome_WidgetWin_1
chromewidgetwin.Wait('ready')
chromewidgetwin.ClickInput()
chromewidgetwin.type_keys ('craig boman')
