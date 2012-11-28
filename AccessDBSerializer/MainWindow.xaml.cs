using GalaSoft.MvvmLight.Messaging;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AccessDBSerializer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public  MainWindowVM _vm;

        public MainWindow()
        {
            InitializeComponent();

            _vm = new MainWindowVM();
            this.DataContext = _vm;

            // Register to receive status updates and show them in the results listbox
            Messenger.Default.Register<Messaging.StatusUpdateMessage>(this, (action => Dispatcher.Invoke(
                (Action)delegate
            {
                this.listResults.Items.Insert(0, DateTime.Now.ToLongTimeString() + ": " + action.MessageText);
            })));
        }

        private void btnDecompose_Click(object sender, RoutedEventArgs e)
        {
            _vm.Decompose();
        }

        private void btnRecompose_Click(object sender, RoutedEventArgs e)
        {
            _vm.Recompose();
        }
    }
}
