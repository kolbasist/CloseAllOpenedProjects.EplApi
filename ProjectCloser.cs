using MyEplanAPI.Service;
using Eplan.EplApi.DataModel;
using System;
using System.Linq;
using System.Threading;

namespace Warden
{
    internal class ProjectCloser
    {
        private ProjectManager _projectManager;
        private SynchronizationContext _synchronizationContext;

        public ProjectCloser()
        {
            _projectManager = new ProjectManager();
            _projectManager.LockProjectByDefault = false;
            _synchronizationContext = SynchronizationContext.Current;
        }

        // Close all projects
        public void CloseAll()
        {
            Logger.SendMSGToEplanLog("[ProjectCloser.CloseAll] Invoked");

            if (SynchronizationContext.Current == _synchronizationContext)
            {
                // Already on UI thread
                ExecuteCloseAll();
            }
            else
            {
                if (_synchronizationContext == null)
                {
                    Logger.SendMSGToEplanLog("[ProjectCloser.CloseAll] SynchronizationContext is null.");
                    return;
                }

                _synchronizationContext.Post(state =>
                {
                    Logger.SendMSGToEplanLog("[ProjectCloser.CloseAll] Switching to UI thread to close projects.");
                    ExecuteCloseAll();
                }, null);
            }
        }

        // Execute close all projects logic
        private void ExecuteCloseAll()
        {
            Logger.SendMSGToEplanLog("[ProjectCloser.ExecuteCloseAll] Attempting to close all projects.");

            try
            {
                while (_projectManager.OpenProjects.Length > 0)
                {
                    _projectManager.CurrentProject.Close();
                    Logger.SendMSGToEplanLog("[ProjectCloser.ExecuteCloseAll] Closed one project. Remaining projects: " + _projectManager.OpenProjects.Length);
                }
            }
            catch (Exception ex)
            {
                Logger.SendMSGToEplanLog($"[ProjectCloser.ExecuteCloseAll] Exception while closing projects: {ex.Message}");
            }
        }

        // Close a specific project by name
        public void CloseProject(string projectName)
        {
            Logger.SendMSGToEplanLog($"[ProjectCloser.CloseProject] Invoked for project: {projectName}");

            if (SynchronizationContext.Current == _synchronizationContext)
            {
                // Already on UI thread
                ExecuteCloseProject(projectName);
            }
            else
            {
                if (_synchronizationContext == null)
                {
                    Logger.SendMSGToEplanLog("[ProjectCloser.CloseProject] SynchronizationContext is null.");
                    return;
                }

                _synchronizationContext.Post(state =>
                {
                    Logger.SendMSGToEplanLog($"[ProjectCloser.CloseProject] Switching to UI thread to close project: {projectName}");
                    ExecuteCloseProject(projectName);
                }, null);
            }
        }

        // Execute close specific project logic
        private void ExecuteCloseProject(string projectName)
        {
            Logger.SendMSGToEplanLog($"[ProjectCloser.ExecuteCloseProject] Attempting to close project: {projectName}");

            try
            {
                Project project = _projectManager.OpenProjects.FirstOrDefault(p => p.ProjectName == projectName);
                if (project != null)
                {
                    project.Close();
                    Logger.SendMSGToEplanLog($"[ProjectCloser.ExecuteCloseProject] Project '{projectName}' closed successfully.");
                }
                else
                {
                    Logger.SendMSGToEplanLog($"[ProjectCloser.ExecuteCloseProject] Project '{projectName}' not found.");
                }
            }
            catch (Exception ex)
            {
                Logger.SendMSGToEplanLog($"[ProjectCloser.ExecuteCloseProject] Exception while closing project '{projectName}': {ex.Message}");
            }
        }       
    }
}
