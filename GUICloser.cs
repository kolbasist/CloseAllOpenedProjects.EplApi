using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.EServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CENTEC.EplAPI.Service;
using Eplan.EplApi.System;
using Eplan.EplApi.HEServices;

namespace CENTEC.EplAddin.TestEplanAPI
{
    internal class GUICloser : IEplAction
    {
        private ActionCallingContext _callingContext;
        private ProjectCloser _closer;

        public GUICloser()
        {
            _closer = new ProjectCloser();
            _callingContext = new ActionCallingContext();
            _callingContext.AddParameter("_cmdline", "CloseAllProjects");
        }

        public bool Execute(ActionCallingContext ctx)
        {
            _closer.CloseAll();
            return true;
        }

        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "CloseAllProjects";
            Ordinal = 20;
            return true;
        }

        public void GetActionProperties(ref ActionProperties actionProperties)
        {
            ActionParameterProperties description = new ActionParameterProperties();
            description.Set("Action test with parameters.");
            actionProperties.AddParameter(description);
        }

        private void Run()
        {
            try
            {
                Edit ed = new Edit();
                ed.ClearSelection();
                Project[] projects = GetProjects(out ProjectManager projectManager);
                projectManager.LockProjectByDefault = false;

                if (projects.Length != 0)
                {
                    Logger.SendMSGToEplanLog($"[ProjectCloser.Run] Trying to close all opened projects");
                    Logger.SendMSGToEplanLog($"[ProjectCloser.Run] Now {projects.Length} projects are opened");

                    using (new Eplan.EplApi.DataModel.LockingStep())
                        //for (int i = 0; i < projects.Length; i++)
                        //{
                        //    string projectName = projects[i].ProjectName;
                        //    ed.SelectProjectInPagesNavigator(projectManager.OpenProjects.Where(p => p.ProjectName == projectName).FirstOrDefault());
                        //    ed.RedrawGed();
                        //    projectManager.OpenProjects.Where(p => p.ProjectName == projectName).FirstOrDefault().Close();
                        //    Logger.SendMSGToEplanLog($"[ProjectCloser.Run] Project {projectName} was succesfully closed at {DateTime.Now}");
                        //}
                        while (GetProjects(out projectManager).Length > 0)
                        {
                            string currentProjectName = projectManager.CurrentProject.ProjectFullName;
                            projectManager.CurrentProject.Close();
                            Logger.SendMSGToEplanLog($"[ProjectCloser.Run] Project {currentProjectName} was succesfully closed at {DateTime.Now}");
                        }
                }
            }
            catch (Exception ex)
            {
                Logger.SendMSGToEplanLog($"Intent execution error: {ex.Message}");
            }

            //Eplan.EplApi.ApplicationFramework.ActionCallingContext acc = new Eplan.EplApi.ApplicationFramework.ActionCallingContext();
            //acc.AddParameter("PROJECT", currentProjectName);
            //String strAction = @"XPrjActionProjectClose";
            //result = new Eplan.EplApi.ApplicationFramework.CommandLineInterpreter().Execute(strAction, acc);

            //try
            //{
            if (GetProjects(out  ProjectManager PM).Length != 0)
            {
                Run();
            }
            //}
            //catch (Exception ex)
            //{
            //    Logger.SendMSGToEplanLog($"Intent execution error: {ex.Message}");
            //}
        }

        private Project[] GetProjects(out ProjectManager projectManager)
        {
            projectManager = new ProjectManager();
            projectManager.LockProjectByDefault = false;

            using (new Eplan.EplApi.DataModel.LockingStep())
                return projectManager.OpenProjects;

        }
    }
}
