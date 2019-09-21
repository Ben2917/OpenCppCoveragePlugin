﻿// OpenCppCoverage is an open source code coverage for C++.
// Copyright (C) 2014 OpenCppCoverage
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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenCppCoverage.VSPackage.Settings;
using OpenCppCoverage.VSPackage.Settings.UI;
using System.Collections.Generic;
using System.IO;

namespace VSPackage_IntegrationTests
{
    [TestClass]
    public class MainSettingTests: TestHelpers
    {
        readonly Dictionary<PathKind, string> paths = new Dictionary<PathKind, string>();

        //---------------------------------------------------------------------
        enum PathKind
        {
            Cobertura,
            Binary,
            Html,
            ConfigFile,
            UnifiedDiff
        }

        //---------------------------------------------------------------------
        [TestCleanup]
        public void TestCleanUp()
        {
            foreach (var kvp in paths)
            {
                var path = kvp.Value;
                if (Directory.Exists(path))
                    Directory.Delete(path, true);
                else
                    File.Delete(path);
            }
        }

        //---------------------------------------------------------------------
        [TestMethod]
        [HostType("VS IDE")]
        public void CheckAllSettings()
        {            
            OpenSolution(CppConsoleApplication);
            var controller = ExecuteOpenCppCoverageCommand();
            var binaryOutput = CreateBinaryOutput(controller);

            RunInUIhread(() => SetMainSettingValue(controller, binaryOutput));
            RunCoverageAndCheckExitCode(controller);
            Assert.IsTrue(File.Exists(paths[PathKind.Binary]));
            Assert.IsTrue(File.Exists(paths[PathKind.Cobertura]));
            Assert.IsTrue(Directory.Exists(paths[PathKind.Html]));
        }

        //---------------------------------------------------------------------
        void SetMainSettingValue(MainSettingController controller, string binaryOutput)
        {
            controller.BasicSettingController.Arguments = "Test";
            Update(controller.FilterSettingController);
            Update(controller.ImportExportSettingController, binaryOutput);
            Update(controller.MiscellaneousSettingController);
        }

        //---------------------------------------------------------------------
        void Update(FilterSettingController controller)
        {
            controller.Settings.AdditionalSourcePatterns.Add(new BindableString("Test"));
            controller.Settings.AdditionalModulePatterns.Add(new BindableString("Test"));
            controller.Settings.ExcludedSourcePatterns.Add(new BindableString("Test"));
            controller.Settings.ExcludedModulePatterns.Add(new BindableString("Test"));

            var unifiedDiff = GetEmptyFile(PathKind.UnifiedDiff);
            controller.Settings.UnifiedDiffs.Add(new FilterSettings.UnifiedDiff
            {
                OptionalRootFolder = Path.GetDirectoryName(unifiedDiff),
                UnifiedDiffPath = unifiedDiff
            });
        }

        //---------------------------------------------------------------------
        void Update(ImportExportSettingController controller, string binaryOutput)
        {
            controller.Settings.Exports.Clear();
            controller.Settings.Exports.Add(new ImportExportSettingController.Export
                {   Type = ImportExportSettings.Type.Cobertura,
                    Path = GetTemporaryPath(PathKind.Cobertura) });
            controller.Settings.Exports.Add(new ImportExportSettingController.Export
                {   Type = ImportExportSettings.Type.Html,
                    Path = GetTemporaryPath(PathKind.Html) });

            controller.Settings.InputCoverages.Add(new BindableString(binaryOutput));
            controller.Settings.CoverChildrenProcesses = !controller.Settings.CoverChildrenProcesses;
            controller.Settings.AggregateByFile = !controller.Settings.AggregateByFile;
        }

        //---------------------------------------------------------------------
        void Update(MiscellaneousSettingController controller)
        {
            controller.Settings.OptionalConfigFile = GetEmptyFile(PathKind.ConfigFile);
            controller.Settings.LogTypeValue = MiscellaneousSettings.LogType.Verbose;
            controller.Settings.ContinueAfterCppExceptions = !controller.Settings.ContinueAfterCppExceptions;        
        }

        //---------------------------------------------------------------------
        string GetTemporaryPath(PathKind kind)
        {
            var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            paths.Add(kind, path);

            return path;
        }

        //---------------------------------------------------------------------
        string GetEmptyFile(PathKind kind)
        {
            var path = Path.GetTempFileName();
            paths.Add(kind, path);

            return path;
        }

        //---------------------------------------------------------------------
        string CreateBinaryOutput(MainSettingController controller)
        {
            string path = GetTemporaryPath(PathKind.Binary);
            RunInUIhread(() =>
            {
                controller.ImportExportSettingController.Settings.Exports.Add(new ImportExportSettingController.Export
                { Type = ImportExportSettings.Type.Binary, Path = path });

                controller.MiscellaneousSettingController.Settings.LogTypeValue = MiscellaneousSettings.LogType.Quiet;                
            });
            RunCoverageAndCheckExitCode(controller);
            return path;
        }

        //---------------------------------------------------------------------
        void RunCoverageAndCheckExitCode(MainSettingController controller)
        {
            var coverageTreeController = RunCoverageCommandAndWait(controller);
            Assert.IsTrue(string.IsNullOrEmpty(coverageTreeController.Warning));
        }
    }
}
