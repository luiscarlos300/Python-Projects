import win32com.client
import sys
import subprocess
import time

class SapGui():
    def __init__(self):
        self.path = r"C:\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(5)
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return
        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection("PRD []", True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()

    def sapLogin(self):
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = ""  # Mandante
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"  # Idioma
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = ""  # Usuario
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = ""  # Senha
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 13
            self.session.findById("wnd[0]").sendVKey(0)

            # Pipeline FBL1N - BI Contas a Pagar e a Receber
            self.session.findById("wnd[0]").resizeWorkingPane(150, 25, False)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = ""
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/chkX_SHBV").selected = True
            self.session.findById("wnd[0]/usr/chkX_MERK").selected = True
            self.session.findById("wnd[0]/usr/chkX_PARK").selected = True
            self.session.findById("wnd[0]/usr/chkX_APAR").selected = True
            self.session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = "T001"
            self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/FBL1N_TELEC"
            self.session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
            self.session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 12
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = "/01.  Bases BI"
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 93
            self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[12]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[12]").press()

        except:
            print(sys.exc_info()[0])

if __name__ == '__main__':
    SapGui().sapLogin()
    subprocess.call("TASKKILL /F /IM saplogon.exe", shell=True)
    subprocess.call("TASKKILL /F /IM excel.exe", shell=True)