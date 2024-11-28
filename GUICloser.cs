using Eplan.EplApi.ApplicationFramework;
using CENTEC.EplAPI.Service;

namespace CENTEC.EplAddin.TestEplanAPI
{
    internal class GUICloser : IEplAction
    {
        private ProjectCloser _closer;

        public GUICloser()
        {
            _closer = new ProjectCloser();
        }

        public bool Execute(ActionCallingContext ctx)
        {
            Logger.SendMSGToEplanLog("[GUICloser.Execute] Invoked to close all projects.");
            _closer.CloseAll();
            return true;
        }

        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "CloseAllProjects";
            Ordinal = 20;
            Logger.SendMSGToEplanLog("[GUICloser.OnRegister] Action 'CloseAllProjects' registered.");
            return true;
        }

        public void GetActionProperties(ref ActionProperties actionProperties)
        {
            ActionParameterProperties description = new ActionParameterProperties();
            description.Set("Action to close all projects.");
            actionProperties.AddParameter(description);
            Logger.SendMSGToEplanLog("[GUICloser.GetActionProperties] Action properties set.");
        }
    }
}
