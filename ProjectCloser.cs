using CENTEC.EplAPI.Service;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CENTEC.EplAddin.TestEplanAPI
{
    internal class ProjectCloser
    {
        private ProjectManager _projectManager;
        private LockingStep _lockingStep;
        private LockingException _lockException;

        public ProjectCloser()
        {
            _projectManager = new ProjectManager();
            _projectManager.LockProjectByDefault = false;
        }

        private void CloseProject(string projectName)
        {
            try
            {
                Edit ed = new Edit();
                ed.ClearSelection();
                _projectManager.LockProjectByDefault = false;

                using (new Eplan.EplApi.DataModel.LockingStep())
                    if (_projectManager.OpenProjects.Any(p => p.ProjectName == projectName))
                    {
                        ed.SelectProjectInPagesNavigator(_projectManager.OpenProjects.Where(p => p.ProjectName == projectName).FirstOrDefault());
                        ed.RedrawGed();

                        //PM.CurrentProject.Close();
                        _projectManager.OpenProjects.Where(p => p.ProjectName == projectName).FirstOrDefault().Close();

                        Logger.SendMSGToEplanLog($"Project '{projectName}' closed");
                    }

                //Close by action (use full project name)
                //Eplan.EplApi.ApplicationFramework.ActionCallingContext acc = new Eplan.EplApi.ApplicationFramework.ActionCallingContext();
                //acc.AddParameter("PROJECT", ProjectName);
                //String strAction = @"XPrjActionProjectClose";
                //result = new Eplan.EplApi.ApplicationFramework.CommandLineInterpreter().Execute(strAction, acc);

            }
            catch (LockingException ex)
            {
                Logger.SendMSGToEplanLog($"(Locking) Intent execution error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Logger.SendMSGToEplanLog($"Intent execution error: {ex.Message}");
            }

            if (_projectManager.OpenProjects.Any(p => p.ProjectName == projectName))
            {
                Logger.SendMSGToEplanLog("[ProjectCloser] Mario, your princess is in another castle...");
                CloseProject(projectName);
            }
        }

        private Queue<string> GetProjectsList()
        {
            _projectManager.LockProjectByDefault = false;
            Project[] projects = _projectManager.OpenProjects;
            Queue<string> projectNames = new Queue<string> { };

            foreach (Project p in projects)             
                projectNames.Enqueue(p.ProjectName);            

            return projectNames;
        }

        public void CloseAll()
        {
            Queue<string> projectNames = GetProjectsList();

            while (projectNames.Count > 0)
            {
                CloseProject(projectNames.Dequeue());
            }
        }

        //public void SelfDestruct()
        //{

        //}
    }
}
