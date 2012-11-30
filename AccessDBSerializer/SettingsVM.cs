using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace AccessDBSerializer
{
    public class SettingsVM : INotifyPropertyChanged
    {
        public string WorkingFolder
        {
            get { return Properties.Settings.Default.WorkingFolder; }
            set
            {
                Properties.Settings.Default.WorkingFolder = value;
                Properties.Settings.Default.Save();
                PropertyChanged(this, new PropertyChangedEventArgs("WorkingFolder"));
            }
        }

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

        internal void ChangeWorkingFolder()
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.Description = "Working Root Folder (location of the access database).  Extracted source files will be placed in a subdirectory named \"Source\".";
            fbd.SelectedPath = Properties.Settings.Default.WorkingFolder;
            System.Windows.Forms.DialogResult result = fbd.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                WorkingFolder = fbd.SelectedPath;
            }
        }

        internal void ChangeAccessFilename()
        {
            System.Windows.Forms.OpenFileDialog fd = new System.Windows.Forms.OpenFileDialog();
            fd.DefaultExt = "accdb";
            fd.InitialDirectory = Properties.Settings.Default.WorkingFolder;
            System.Windows.Forms.DialogResult result = fd.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                AccessDBFilename = fd.FileName;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
