from pywinauto.application import Application

app = Application().Connect(title=u'Sierra \xb7 Miami University Libraries \xb7 Craig Boman', class_name='SunAwtFrame')
sunawtframe = app.SunAwtFrame
sunawtframe.SetFocus()
sunawtframe.ClickInput()
sunawtframe.Edit.Typekeys ('craig boman')