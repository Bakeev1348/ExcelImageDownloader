using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace ExcelImageDownloader
{
    //class needed to communicate between ThisAddin class and Panel
    //Panel's fillAddress method subscribes to This.hasAddress event
    public class sender
    {
        private string _address;
        private bool _enabled;
        private bool _add;
        private RibbonEditBox _editBox;
        private RibbonToggleButton _toggleButton;

        public delegate void voidMethod();
        public event voidMethod hasAddress;

        public sender()
        {
            _address = null;
            _editBox = null;
            _toggleButton = null;
            _enabled = false;
            _add = false;
        }
        public void enable(RibbonEditBox control, RibbonToggleButton toggleButton, bool add)
        {
            _enabled = true;
            _add = add;
            _toggleButton = toggleButton;
            _editBox = control;
        }
        public bool enabled()
        {
            return _enabled;
        }
        public bool adding()
        {
            return _add;
        }
        public void disable()
        {
            _address = null;
            _editBox = null;
            _toggleButton = null;
            _enabled = false;
            _add = false;
        }

        //hasAddress event HERE
        public void setAddress(string address)
        {
            _address = address;
            hasAddress();
        }
        public string getAddress()
        {
            return _address;
        }
        public RibbonEditBox getEditBox()
        {
            return _editBox;
        }
        public RibbonToggleButton getToggleButton()
        {
            return _toggleButton;
        }
    }

    public partial class ThisAddIn
    {
        //contains variables that contain current app, current book and current sheet
        public static Excel.Worksheet activeWorksheet;
        public static Excel.Workbook thisWorkbook;
        public static Excel.Application thisApp;
        public static sender sender;

        //initialises thisWorkbook and activeWorksheet
        private void Application_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            try
            {
                ThisAddIn.activeWorksheet = (Excel.Worksheet)Application.ActiveSheet;
                ThisAddIn.thisWorkbook = Wb;
                sender = new sender();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //initialises activeWorksheet
        private void Application_SheetActivate(object Sh)
        {
            try
            {
                ThisAddIn.activeWorksheet = (Excel.Worksheet)Application.ActiveSheet;
                ThisAddIn.thisWorkbook = (Excel.Workbook)Application.ActiveWorkbook;
                sender = new sender();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //calls this.sender.setAddress that calls hasAddress event
        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            if (sender.enabled())
            {
                sender.setAddress(Target.Address);
            }
        }

        //initialises thisApp and subscribes methods to app events
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);
            this.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Application_SheetActivate);
            this.Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            thisApp = Application;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }



}
