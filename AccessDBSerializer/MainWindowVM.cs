using GalaSoft.MvvmLight.Messaging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using AccessDBSerializer.Messaging;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Collections;
using System.ComponentModel;

namespace AccessDBSerializer
{
    public class MainWindowVM : INotifyPropertyChanged
    {

        [DllImport("ole32.dll")]
        static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        public string AccessDBFilename
        {
            get { return Properties.Settings.Default.AccessDBFilename; }
            set
            {
                Properties.Settings.Default.AccessDBFilename = value;
                Properties.Settings.Default.Save();
                PropertyChanged(this, new PropertyChangedEventArgs("AccessDBFilename"));
            }
        }

        internal void Decompose()
        {

            Task.Factory.StartNew(() =>
                {
                    string exportPath = Properties.Settings.Default.WorkingFolder;
                    string accessDatabaseFilename =  Properties.Settings.Default.AccessDBFilename;

                    ExportModulesTxt(accessDatabaseFilename, exportPath);
                }
            );
        }


        internal void Recompose()
        {
            Task.Factory.StartNew(() =>
                {
                    string importPath = Properties.Settings.Default.WorkingFolder;
                    string accessDatabaseFilename = Properties.Settings.Default.AccessDBFilename;

                    ImportModulesText(accessDatabaseFilename, importPath);
                }
            );
        }

        private void ImportModulesText(string accessDatabaseFilename, string importPath)
        {
            try
            {

                FileInfo fi = new FileInfo(accessDatabaseFilename);
                if (importPath == "")
                {
                    importPath = fi.Directory.ToString() + @"\Source\";
                }
                else
                {
                    importPath = importPath + @"\Source\";
                }
                string stubADPFilename = importPath + fi.Name.Replace(fi.Extension, "") + "_stub" + fi.Extension;

                // Back up then replace the base file with the stub
                if (fi.Exists)
                {
                    File.Copy(accessDatabaseFilename, accessDatabaseFilename + ".bak", true);
                }

                File.Copy(stubADPFilename, accessDatabaseFilename, true);

                // Launch Access
                PublishStatusMessage("Starting Access");

                Type t = null;
                object app = CoCreate("Access.Application", ref t);

                if (app != null)
                {
                    t.InvokeMember("OpenCurrentDatabase", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { accessDatabaseFilename });
                    t.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });

                    PublishStatusMessage("Successfully Opened Access Application");

                    DirectoryInfo folder = new DirectoryInfo(importPath);

                    foreach (var file in folder.EnumerateFiles())
                    {
                        string objectType = file.Extension.Substring(1);
                        string objectName = file.Name.Substring(0, (file.Name.Length - file.Extension.Length));
                        PublishStatusMessage("Importing " + objectType + " " + objectName);

                        switch (objectType)
                        {
                            case "form":
                                t.InvokeMember("LoadFromText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acForm, objectName, importPath + @"\" + objectName + ".form" });
                                break;
                            case "bas":
                                t.InvokeMember("LoadFromText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acModule, objectName, importPath + @"\" + objectName + ".bas" });
                                break;
                            case "mac":
                                t.InvokeMember("LoadFromText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acMacro, objectName, importPath + @"\" + objectName + ".mac" });
                                break;
                            case "report":
                                t.InvokeMember("LoadFromText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acReport, objectName, importPath + @"\" + objectName + ".report" });
                                break;
                            default:
                                break;
                        }
                    }

                    PublishStatusMessage("Calling Command acCmdCompileAndSaveAllModules");
                    t.InvokeMember("RunCommand", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acCmdCompileAndSaveAllModules });

                    PublishStatusMessage("Closing Database...");
                    t.InvokeMember("CloseCurrentDatabase", System.Reflection.BindingFlags.InvokeMethod, null, app, null);

                    PublishStatusMessage("Calling CompactRepair...");
                    t.InvokeMember("CompactRepair", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { accessDatabaseFilename, accessDatabaseFilename + "_" });

                    PublishStatusMessage("Calling Quit...");
                    t.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, null);

                    PublishStatusMessage("Cleaning up temp files...");
                    File.Copy(accessDatabaseFilename + "_", accessDatabaseFilename, true);
                    File.Delete(accessDatabaseFilename + "_");

                    PublishStatusMessage("Finished");

                }
            }
            catch (Exception ex)
            {
                PublishStatusMessage("EXCEPTION: " + ex.Message);
            }
        }

        private void ExportModulesTxt(string accessDatabaseFilename, string exportPath)
        {
            try
            {
                FileInfo fi = new FileInfo(accessDatabaseFilename);

                if (exportPath == "")
                {
                    exportPath = fi.Directory.ToString() + @"\Source\";
                }
                else
                {
                    exportPath = exportPath + @"\Source\";
                }

                string stubADPFilename = exportPath + fi.Name.Replace(fi.Extension, "") + "_stub" + fi.Extension;

                PublishStatusMessage("copy stub to " + stubADPFilename + "...");

                Trace.Write("copy stub to " + stubADPFilename + "...");

                if (!Directory.Exists(exportPath))
                {
                    PublishStatusMessage("creating directory " + exportPath);
                    Messenger.Default.Send<StatusUpdateMessage>(new StatusUpdateMessage() { MessageText = "creating directory " + exportPath });
                    Directory.CreateDirectory(exportPath);
                }
                File.Copy(accessDatabaseFilename, stubADPFilename, true);

                Type t = null;
                object app = CoCreate("Access.Application", ref t);

                if (app != null)
                {
                    object doCmd = t.InvokeMember("DoCmd", System.Reflection.BindingFlags.GetProperty, null, app, null);
                    Type doCmdType = doCmd.GetType();

                    t.InvokeMember("OpenCurrentDatabase", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { stubADPFilename });
                    t.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });

                    PublishStatusMessage("Successfully Opened Access Application");

                    Dictionary<string, object> dctDelete = new Dictionary<string, object>();

                    PublishStatusMessage("Exporting...");

                    object currentProject = null;
                    object allForms = null;

                    currentProject = t.InvokeMember("CurrentProject", System.Reflection.BindingFlags.GetProperty, null, app, null);
                    Type currentProjectType = currentProject.GetType();

                    // Process Forms
                    allForms = currentProjectType.InvokeMember("AllForms", System.Reflection.BindingFlags.GetProperty, null, currentProject, null);
                    Type allFormsType = allForms.GetType();
                    Type formType = null;
                    foreach (var item in (IEnumerable)allForms)
                    {
                        if (formType == null)
                        {
                            formType = item.GetType();
                        }
                        string formName = (string)formType.InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, item, null);
                        PublishStatusMessage("Serializing Form: " + formName);

                        t.InvokeMember("SaveAsText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acForm, formName, exportPath + @"\" + formName + ".form" });
                        doCmdType.InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, doCmd, new object[] { AccessObjectType.acForm, formName });
                        dctDelete.Add("FO" + formName, AccessObjectType.acForm);
                    }

                    // Process Modules
                    object allModules = null;
                    allModules = currentProjectType.InvokeMember("AllModules", System.Reflection.BindingFlags.GetProperty, null, currentProject, null);
                    Type allModulesType = allModules.GetType();
                    Type moduleType = null;
                    foreach (var item in (IEnumerable)allModules)
                    {
                        if (moduleType == null)
                        {
                            moduleType = item.GetType();
                        }
                        string moduleName = (string)moduleType.InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, item, null);
                        PublishStatusMessage("Serializing Module: " + moduleName);

                        t.InvokeMember("SaveAsText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acModule, moduleName, exportPath + @"\" + moduleName + ".bas" });
                        dctDelete.Add("MO" + moduleName, AccessObjectType.acModule);
                    }

                    // Process Macros
                    object allMacros = null;
                    allMacros = currentProjectType.InvokeMember("AllMacros", System.Reflection.BindingFlags.GetProperty, null, currentProject, null);
                    Type allMacrosType = allModules.GetType();
                    Type macroType = null;
                    foreach (var item in (IEnumerable)allMacros)
                    {
                        if (macroType == null)
                        {
                            macroType = item.GetType();
                        }
                        string macroName = (string)macroType.InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, item, null);
                        PublishStatusMessage("Serializing Macro: " + macroName);

                        t.InvokeMember("SaveAsText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acMacro, macroName, exportPath + @"\" + macroName + ".mac" });
                        dctDelete.Add("MA" + macroName, AccessObjectType.acMacro);
                    }

                    // Process Reports
                    object allReports = null;
                    allReports = currentProjectType.InvokeMember("AllReports", System.Reflection.BindingFlags.GetProperty, null, currentProject, null);
                    Type allReportsType = allModules.GetType();
                    Type reportType = null;
                    foreach (var item in (IEnumerable)allReports)
                    {
                        if (reportType == null)
                        {
                            reportType = item.GetType();
                        }
                        string reportName = (string)macroType.InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, item, null);
                        PublishStatusMessage("Serializing Report: " + reportName);

                        t.InvokeMember("SaveAsText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acReport, reportName, exportPath + @"\" + reportName + ".report" });
                        dctDelete.Add("RE" + reportName, AccessObjectType.acReport);
                    }

                    // Clean out the database so it can be compacted
                    PublishStatusMessage("Deleting...");
                    foreach (var item in dctDelete)
                    {
                        PublishStatusMessage("Deleting " + item.Key);
                        doCmdType.InvokeMember("DeleteObject", System.Reflection.BindingFlags.InvokeMethod, null, doCmd, new object[] { item.Value, item.Key.Substring(2) });
                    }

                    // Cleanup
                    PublishStatusMessage("Closing Database...");
                    t.InvokeMember("CloseCurrentDatabase", System.Reflection.BindingFlags.InvokeMethod, null, app, null);

                    PublishStatusMessage("Calling CompactRepair...");
                    t.InvokeMember("CompactRepair", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { stubADPFilename, stubADPFilename + "_" });

                    PublishStatusMessage("Calling Quit...");
                    t.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, null);

                    PublishStatusMessage("Cleaning up temp files...");
                    File.Copy(stubADPFilename + "_", stubADPFilename, true);
                    File.Delete(stubADPFilename + "_");

                    PublishStatusMessage("Finished");
                }
            }
            catch (Exception ex)
            {
                PublishStatusMessage("EXCEPTION: " + ex.Message);
            }
        }


        internal void PublishStatusMessage(string message)
        {
            Dispatcher.CurrentDispatcher.Invoke((Action)delegate { Messenger.Default.Send<StatusUpdateMessage>(new StatusUpdateMessage() { MessageText = message }); }, null);
        }

        public object CoCreate(string name, ref Type t)
        {
            Guid newclsid;
            object idisp;
            CLSIDFromProgID(name, out newclsid);  // get clsid from com object  
            t = Type.GetTypeFromCLSID(newclsid);  // received the com objects type  
            idisp = Activator.CreateInstance(t);  // create an object with the given type    
            return idisp;
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
