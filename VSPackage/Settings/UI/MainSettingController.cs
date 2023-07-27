// OpenCppCoverage is an open source code coverage for C++.
// Copyright (C) 2016 OpenCppCoverage
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using EnvDTE;
using EnvDTE80;
using GalaSoft.MvvmLight.Command;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.VisualStudio.Shell;
using OpenCppCoverage.VSPackage.Helper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Controls;
using System.Windows.Input;

namespace OpenCppCoverage.VSPackage.Settings.UI
{
    //-------------------------------------------------------------------------
    class MainSettingController : PropertyChangedNotifier
    {
        readonly IOpenCppCoverageCmdLine openCppCoverageCmdLine;
        readonly ISettingsStorage settingsStorage;
        readonly CoverageRunner coverageRunner;
        readonly IServiceProvider serviceProvider;

        string selectedProjectPath;
        string solutionConfigurationName;
        bool displayProgramOutput;
        ProjectSelectionKind kind;

        //---------------------------------------------------------------------
        public MainSettingController(
            ISettingsStorage settingsStorage,
            IOpenCppCoverageCmdLine openCppCoverageCmdLine,
            CoverageRunner coverageRunner,
            IServiceProvider serviceProvider)
        {
            this.settingsStorage = settingsStorage;
            this.openCppCoverageCmdLine = openCppCoverageCmdLine;
            this.RunCoverageCommand = new RelayCommand(() => OnRunCoverageCommand());
            this.CloseCommand = new RelayCommand(() =>
            {
                this.CloseWindowEvent?.Invoke(this, EventArgs.Empty);
            });
            this.ResetToDefaultCommand = new RelayCommand(
                () => UpdateStartUpProject(ComputeStartUpProjectSettings(kind)));
            this.BasicSettingController = new BasicSettingController();
            this.FilterSettingController = new FilterSettingController();
            this.ImportExportSettingController = new ImportExportSettingController();
            this.MiscellaneousSettingController = new MiscellaneousSettingController();

            this.coverageRunner = coverageRunner;
            this.serviceProvider = serviceProvider;
        }

        //---------------------------------------------------------------------
        public void UpdateFields(ProjectSelectionKind kind, bool displayProgramOutput)
        {
            var settings = ComputeStartUpProjectSettings(kind);
            this.UpdateStartUpProject(settings);
            this.selectedProjectPath = settings.ProjectPath;
            this.displayProgramOutput = displayProgramOutput;
            this.solutionConfigurationName = settings.SolutionConfigurationName;
            this.kind = kind;

            var uiSettings = this.settingsStorage.TryLoad(this.selectedProjectPath, this.solutionConfigurationName);

            if (uiSettings != null)
            {
                this.BasicSettingController.UpdateSettings(uiSettings.BasicSettingController);
                this.FilterSettingController.UpdateSettings(uiSettings.FilterSettingController);
                this.ImportExportSettingController.UpdateSettings(uiSettings.ImportExportSettingController);
                this.MiscellaneousSettingController.UpdateSettings(uiSettings.MiscellaneousSettingController);
            }
        }

        List<ExtendedProject> CreateExtendedProjectsFor(Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var projects = new List<ExtendedProject>();

            if (project.Kind == EnvDTE80.ProjectKinds.vsProjectKindSolutionFolder)
            {
                foreach (ProjectItem projectItem in project.ProjectItems)
                {
                    var subProject = projectItem.SubProject;
                    if (subProject != null)
                        projects.AddRange(CreateExtendedProjectsFor(subProject));
                }
            }
            else
            {
                dynamic projectObject = project.Object;

                try
                {
                    if (projectObject != null && projectObject.Kind == "VCProject")
                        projects.Add(new ExtendedProject(project, new DynamicVCProject(projectObject)));
                }
                catch (RuntimeBinderException)
                {
                    // Nothing because not a VCProject
                }
            }

            return projects;
        }

        //---------------------------------------------------------------------
        StartUpProjectSettings ComputeStartUpProjectSettings(ProjectSelectionKind kind)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = (DTE2)serviceProvider.GetService(typeof(EnvDTE.DTE));

            var solution = (Solution2)dte.Solution;
            var solutionBuild = (SolutionBuild2)solution.SolutionBuild;
            var activeConfiguration = (SolutionConfiguration2)solutionBuild.ActiveConfiguration;

            if (activeConfiguration != null)
            {
                var projects = new List<ExtendedProject>();
                foreach (Project pj in solution.Projects)
                {
                    projects.AddRange(CreateExtendedProjectsFor(pj));
                }

                ExtendedProject project = null;
                switch (kind)
                {
                    case ProjectSelectionKind.StartUpProject:
                        var startupProjectsNames = solution.SolutionBuild.StartupProjects as object[];

                        if (startupProjectsNames == null)
                            return null;

                        var startupProjectsSet = new HashSet<String>();
                        foreach (String pjName in startupProjectsNames)
                            startupProjectsSet.Add(pjName);

                        project = projects.Where(p => startupProjectsSet.Contains(p.UniqueName)).FirstOrDefault();
                        break;
                    case ProjectSelectionKind.SelectedProject:
                        var selectedProjects = ((Array)dte.ActiveSolutionProjects).Cast<Project>();

                        if (selectedProjects.Count() != 1)
                            return null;

                        var projectName = selectedProjects.First().UniqueName;
                        project = projects.Where(p => p.UniqueName == projectName).FirstOrDefault();
                        break;
                }

                if (project == null)
                    goto Cleanup_and_exit;

                DynamicVCConfiguration startupConfiguration = null;
                var contexts = activeConfiguration.SolutionContexts.Cast<SolutionContext>();
                var context = contexts.FirstOrDefault(c =>
                {
                    Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
                    return c.ProjectName == project.UniqueName;
                });

                if (context == null)
                    startupConfiguration = null;

                if (!context.ShouldBuild)
                {
                    startupConfiguration = null;
                }

                var configurations = project.Configurations;
                startupConfiguration = configurations.FirstOrDefault(
                    c =>
                    {
                        Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
                        return c.ConfigurationName == context.ConfigurationName && c.PlatformName == context.PlatformName;
                    });

                var debugSettings = startupConfiguration.DebugSettings;

                var settings = new StartUpProjectSettings();
                settings.WorkingDir = startupConfiguration.Evaluate(debugSettings.WorkingDirectory);
                settings.Arguments = startupConfiguration.Evaluate(debugSettings.CommandArguments);
                settings.Command = startupConfiguration.Evaluate(debugSettings.Command);
                settings.SolutionConfigurationName = activeConfiguration.Name + '|' + activeConfiguration.PlatformName;
                settings.ProjectName = project.UniqueName;
                settings.ProjectPath = project.Path;

                var cppProjects = new List<StartUpProjectSettings.CppProject>();

                int projectsProcessed = 0;
                foreach (var pj in projects)
                {
                    var localContext = contexts.FirstOrDefault(c =>
                    {
                        Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
                        return c.ProjectName == project.UniqueName;
                    });

                    if (localContext == null)
                        startupConfiguration = null;

                    if (!localContext.ShouldBuild)
                    {
                        startupConfiguration = null;
                    }

                    var localConfigurations = project.Configurations;
                    var localConfiguration = localConfigurations.FirstOrDefault(
                        c =>
                        {
                            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
                            return c.ConfigurationName == localContext.ConfigurationName && c.PlatformName == localContext.PlatformName;
                        });

                    if (localConfiguration != null)
                    {
                        var cppProject = new StartUpProjectSettings.CppProject()
                        {
                            ModulePath = localConfiguration.PrimaryOutput,
                            SourcePaths = PathHelper.ComputeCommonFolders(pj.Files.Select(f => f.FullPath)),
                            Path = pj.UniqueName
                        };
                        cppProjects.Add(cppProject);
                    }

                    projectsProcessed++;
                    Debug.WriteLine("Processed project {0} of {1}", projectsProcessed, projects.Count());
                }

                settings.CppProjects = cppProjects;

                settings.IsOptimizedBuildEnabled = false;

                var environmentVariables = new List<KeyValuePair<string, string>>();
                string environmentStr = startupConfiguration.Evaluate("$(LocalDebuggerEnvironment)");

                foreach (var str in environmentStr.Split('\n'))
                {
                    var equalIndex = str.IndexOf('=');
                    if (equalIndex != -1 && equalIndex != str.Length - 1)
                    {
                        var key = str.Substring(0, equalIndex);
                        var value = str.Substring(equalIndex + 1);

                        environmentVariables.Add(new KeyValuePair<string, string>(key, value));
                    }
                }

                settings.EnvironmentVariables = environmentVariables;

                if (settings != null)
                    return settings;
            }

Cleanup_and_exit:
            Marshal.ReleaseComObject(dte);
            return new StartUpProjectSettings
            {
                CppProjects = new List<StartUpProjectSettings.CppProject>()
            };
        }

        //---------------------------------------------------------------------
        void UpdateStartUpProject(StartUpProjectSettings settings)
        {
            this.BasicSettingController.UpdateStartUpProject(settings);
            this.FilterSettingController.UpdateStartUpProject();
            this.ImportExportSettingController.UpdateStartUpProject();
            this.MiscellaneousSettingController.UpdateStartUpProject();
        }

        //---------------------------------------------------------------------
        public void SaveSettings()
        {
            var uiSettings = new UserInterfaceSettings
            {
                BasicSettingController = this.BasicSettingController.BuildJsonSettings(),
                FilterSettingController = this.FilterSettingController.Settings,
                ImportExportSettingController = this.ImportExportSettingController.Settings,
                MiscellaneousSettingController = this.MiscellaneousSettingController.Settings
            };
            this.settingsStorage.Save(this.selectedProjectPath, this.solutionConfigurationName, uiSettings);
        }

        //---------------------------------------------------------------------
        public MainSettings GetMainSettings()
        {
            return new MainSettings
            {
                BasicSettings = this.BasicSettingController.GetSettings(),
                FilterSettings = this.FilterSettingController.GetSettings(),
                ImportExportSettings = this.ImportExportSettingController.GetSettings(),
                MiscellaneousSettings = this.MiscellaneousSettingController.GetSettings(),
                DisplayProgramOutput = this.displayProgramOutput
            };
        }

        //---------------------------------------------------------------------
        public BasicSettingController BasicSettingController { get; }
        public FilterSettingController FilterSettingController { get; }
        public ImportExportSettingController ImportExportSettingController { get; }
        public MiscellaneousSettingController MiscellaneousSettingController { get; }

        //---------------------------------------------------------------------
        string commandLineText;
        public string CommandLineText
        {
            get { return this.commandLineText; }
            private set { this.SetField(ref this.commandLineText, value); }
        }

        //---------------------------------------------------------------------
        public static string CommandLineHeader = "Command line";

        public TabItem SelectedTab
        {
            set
            {
                if (value != null && (string)value.Header == CommandLineHeader)
                {
                    try
                    {
                        this.CommandLineText = this.openCppCoverageCmdLine.Build(this.GetMainSettings(), "\n");
                    } 
                    catch (Exception e)
                    {
                        this.CommandLineText = e.Message;
                    }
                }
            }
        }
        //---------------------------------------------------------------------
        void OnRunCoverageCommand()
        {
            this.coverageRunner.RunCoverageOnStartupProject(this.GetMainSettings());
        }

        //---------------------------------------------------------------------
        public EventHandler CloseWindowEvent;

        //---------------------------------------------------------------------
        public ICommand CloseCommand { get; }
        public ICommand RunCoverageCommand { get; }
        public ICommand ResetToDefaultCommand { get; }
    }
}
