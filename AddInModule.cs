using System;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
using CENTEC.EplAPI.Service;

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
            var ribbonBar = new RibbonBar(); 
            MultiLangString tabName = new MultiLangString();
            tabName.AddString(ISOCode.Language.L_en_US, "Home");
            tabName.AddString(ISOCode.Language.L_ru_RU, "Главная");
            var tab = ribbonBar.GetTab( tabName, true) ?? ribbonBar.AddTab(tabName);

            try
            {
                string groupName = "CT.API";
                var commandGroup = tab.GetCommandGroup(groupName) ?? tab.AddCommandGroup(groupName);  
                MultiLangString closeButtonText = new MultiLangString();
                closeButtonText.AddString(ISOCode.Language.L_ru_RU,"Закрыть все проекты" );
                closeButtonText.AddString(ISOCode.Language.L_en_US, "Close all Projects");
                closeButtonText.AddString(ISOCode.Language.L_de_DE, "Alle Projekte schliessen");
                commandGroup.AddCommand(closeButtonText, "CloseAllProjects");
            }
            catch (Exception ex)
            {
                Logger.SendMSGToEplanLog($"Ошибка при добавлении команды: {ex.Message}");
            }
            return true;
        }

        public bool OnExit()
        {
            return true;
        }        
    }    
}
