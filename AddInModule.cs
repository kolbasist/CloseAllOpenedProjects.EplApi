using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;
using Microsoft.Win32;
using System.Timers;

namespace CENTEC.EplAddin.TestEplanAPI
{
    public class AddInModule : Eplan.EplApi.ApplicationFramework.IEplAddIn
    {
        private GUICloser _guiCloser;
        private TimersManager _timersManager;

        public bool OnRegister(ref System.Boolean bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }
        public bool OnUnregister()
        {
            return true;
        }
        public bool OnInit()
        {       
            _guiCloser = new GUICloser();
            _timersManager = new TimersManager();
            return true;
        }
        public bool OnInitGui()
        {
            Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
            oMenu.AddMenuItem("Close all", "CloseAllProjects", "Close all opened projects", 35340U, 1, false, true);            
            return true;
        }
        public bool OnExit()
        {
            return true;
        }        
    }    
}
