from pywinauto.application import Application
app = Application(backend="uia").start("Notepad.exe")
app.UntitledNotepad.type_keys('Hello World')
app.UntitledNotepad.type_keys('this is a test')
app.UntitledNotepad.type_keys("%FX")
