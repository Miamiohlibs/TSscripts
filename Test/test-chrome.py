from pywinauto.application import Application

app = Application().Connect(title=u'Sierra \xb7 Miami University Libraries \xb7 Craig Boman - Google Chrome', class_name='Chrome_WidgetWin_1')
chromewidgetwin = app.Chrome_WidgetWin_1
chromewidgetwin.ClickInput()   	