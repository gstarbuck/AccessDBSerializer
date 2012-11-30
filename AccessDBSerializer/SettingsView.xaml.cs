using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AccessDBSerializer
{
    /// <summary>
    /// Interaction logic for SettingsView.xaml
    /// </summary>
    public partial class SettingsView : Window
    {
        private SettingsVM _vm;

        public SettingsView()
        {
            InitializeComponent();
            _vm = new SettingsVM();
            this.DataContext = _vm;
        }

        private void btnChangeWorkingFolder_Click(object sender, RoutedEventArgs e)
        {
            _vm.ChangeWorkingFolder();
        }

        private void btnOk_ChangeAccessFile_Click(object sender, RoutedEventArgs e)
        {
            _vm.ChangeAccessFilename();
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            // Save to pick up any hand-edited values
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
