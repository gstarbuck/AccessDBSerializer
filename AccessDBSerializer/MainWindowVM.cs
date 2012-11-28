using GalaSoft.MvvmLight.Messaging;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using AccessDBSerializer.Messaging;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Collections;

namespace AccessDBSerializer
{
    public class MainWindowVM
    {

        [DllImport("ole32.dll")]
        static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        internal void Decompose()
        {

            Task.Factory.StartNew(() =>
                {
                    String exportPath = "";
                    String accessDatabaseFilename = @"C:\Users\gstarbuck\Documents\Visual Studio 2012\Projects\AccessDBSerializer\AccessDBSerializer\Files\MetrogroDB_new1.accdb";

                    ExportModulesTxt(accessDatabaseFilename, exportPath);
                }
            );
        }

        public enum AccessObjectType
        {
            acDefault = -1,
            acDiagram = 8,
            acForm = 2,
            acFunction = 10,
            acMacro = 4,
            acModule = 5,
            acQuery = 1,
            acReport = 3,
            acServerView = 7,
            acStoredProcedure = 9,
            acTable = 0
        }

        private void ExportModulesTxt(string accessDatabaseFilename, string exportPath)
        {//Function exportModulesTxt(sADPFilename, sExportpath)
            //    Dim myComponent
            //    Dim sModuleType
            //    Dim sTempname
            //    Dim sOutstring

            //    dim myType, myName, myPath, sStubADPFilename
            //    myType = fso.GetExtensionName(sADPFilename)
            //    myName = fso.GetBaseName(sADPFilename)
            //    myPath = fso.GetParentFolderName(sADPFilename)

            FileInfo fi = new FileInfo(accessDatabaseFilename);

            if (exportPath == "")
            {
                exportPath = fi.Directory.ToString() + @"\Source\";
            }

            //    If (sExportpath = "") then
            //        sExportpath = myPath & "\Source\"
            //    End If
            //    sStubADPFilename = sExportpath & myName & "_stub." & myType

            String stubADPFilename = exportPath + fi.Name.Replace(fi.Extension, "") + "_stub" + fi.Extension;

            //'    WScript.Echo "copy stub to " & sStubADPFilename & "..."
            PublishStatusMessage("copy stub to " + stubADPFilename + "...");

            Trace.Write("copy stub to " + stubADPFilename + "...");
            
            //    On Error Resume Next
            //        fso.CreateFolder(sExportpath)
            //    On Error Goto 0
            //    fso.CopyFile sADPFilename, sStubADPFilename

            if (!Directory.Exists(exportPath))
            {
                PublishStatusMessage("creating directory " + exportPath);
                Messenger.Default.Send<StatusUpdateMessage>(new StatusUpdateMessage() { MessageText = "creating directory " + exportPath });
                Directory.CreateDirectory(exportPath);
            }
            File.Copy(accessDatabaseFilename, stubADPFilename, true);

            //'    WScript.Echo "starting Access..."
            //    Dim oApplication
            //    Set oApplication = CreateObject("Access.Application")
            //'    WScript.Echo "opening " & sStubADPFilename & " ..."
            //    If (Right(sStubADPFilename,4) = ".adp") Then
            //        oApplication.OpenAccessProject sStubADPFilename
            //    Else
            //        oApplication.OpenCurrentDatabase sStubADPFilename
            //    End If
            //    oApplication.Visible = false


            Type t = null;
            Object app = CoCreate("Access.Application", ref t);

            if (app != null)
            {
                object doCmd = t.InvokeMember("DoCmd", System.Reflection.BindingFlags.GetProperty, null, app, null);
                Type doCmdType = doCmd.GetType();

                t.InvokeMember("OpenCurrentDatabase", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { stubADPFilename });
                t.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });

                PublishStatusMessage("Successfully Opened Access Application");

                //    dim dctDelete
                //    Set dctDelete = CreateObject("Scripting.Dictionary")

                Dictionary<string, object> dctDelete = new Dictionary<string, object>();

                //'    WScript.Echo "exporting..."
                PublishStatusMessage("Exporting...");
                
                //    Dim myObj
                //    For Each myObj In oApplication.CurrentProject.AllForms
                //'        WScript.Echo "  " & myObj.fullname
                //        oApplication.SaveAsText acForm, myObj.fullname, sExportpath & "\" & myObj.fullname & ".form"
                //        oApplication.DoCmd.Close acForm, myObj.fullname
                //        dctDelete.Add "FO" & myObj.fullname, acForm
                //    Next

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


                //    For Each myObj In oApplication.CurrentProject.AllModules
                //'        WScript.Echo "  " & myObj.fullname
                //        oApplication.SaveAsText acModule, myObj.fullname, sExportpath & "\" & myObj.fullname & ".bas"
                //        dctDelete.Add "MO" & myObj.fullname, acModule
                //    Next

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

                //    For Each myObj In oApplication.CurrentProject.AllMacros
                //'        WScript.Echo "  " & myObj.fullname
                //        oApplication.SaveAsText acMacro, myObj.fullname, sExportpath & "\" & myObj.fullname & ".mac"
                //        dctDelete.Add "MA" & myObj.fullname, acMacro
                //    Next

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

                    t.InvokeMember("SaveAsText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acMacro, macroName, exportPath + @"\" + macroName + ".bas" });
                    dctDelete.Add("MA" + macroName, AccessObjectType.acMacro);
                }

                //    For Each myObj In oApplication.CurrentProject.AllReports
                //'        WScript.Echo "  " & myObj.fullname
                //        oApplication.SaveAsText acReport, myObj.fullname, sExportpath & "\" & myObj.fullname & ".report"
                //        dctDelete.Add "RE" & myObj.fullname, acReport
                //    Next

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

                    t.InvokeMember("SaveAsText", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { AccessObjectType.acReport, reportName, exportPath + @"\" + reportName + ".bas" });
                    dctDelete.Add("RE" + reportName, AccessObjectType.acReport);
                }

                //'    WScript.Echo "deleting..."
                //    dim sObjectname
                //    For Each sObjectname In dctDelete
                //'        WScript.Echo "  " & Mid(sObjectname, 3)
                //        oApplication.DoCmd.DeleteObject dctDelete(sObjectname), Mid(sObjectname, 3)
                //    Next
                PublishStatusMessage("Deleting...");
                foreach (var item in dctDelete)
                {
                    PublishStatusMessage("Deleting " + item.Key);
                    doCmdType.InvokeMember("DeleteObject", System.Reflection.BindingFlags.InvokeMethod, null, doCmd, new object[] { item.Value, item.Key.Substring(2) });
                }

                //    oApplication.CloseCurrentDatabase
                //    oApplication.CompactRepair sStubADPFilename, sStubADPFilename & "_"
                //    oApplication.Quit

                PublishStatusMessage("Closing Database...");
                t.InvokeMember("CloseCurrentDatabase", System.Reflection.BindingFlags.InvokeMethod, null, app, null);

                PublishStatusMessage("Calling CompactRepair...");
                t.InvokeMember("CompactRepair", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { stubADPFilename, stubADPFilename + "_" });

                PublishStatusMessage("Calling Quit...");
                t.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, null);

                //    fso.CopyFile sStubADPFilename & "_", sStubADPFilename
                //    fso.DeleteFile sStubADPFilename & "_"

                PublishStatusMessage("Cleaning up temp files...");
                File.Copy(stubADPFilename + "_", stubADPFilename, true);
                File.Delete(stubADPFilename + "_");

                //    WScript.Echo "Finished"
                PublishStatusMessage("Finished");
                //End Function

            }
        }

        internal void Recompose()
        {
            throw new NotImplementedException();
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
    }
}
