using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Contrôleur_RIB
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Content = new ControleurRIBPage();
        }
        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            ControleurRIBPage controleurRIBPage = MainFrame.Content as ControleurRIBPage;// Allows to call the function to release the Excel file from use when closing the main window (which shuts down the application in the current setup)
            VMControleurRIB vMControleurRIB = controleurRIBPage.DataContext as VMControleurRIB;
            if (vMControleurRIB.ExcelApp != null)
            {
                if (vMControleurRIB.ExcelApp.IsOpen)
                {
                    vMControleurRIB.ExcelApp.Terminate();// Automatically releases the previous file when the user loads a new one without releasing first.
                }
            }
        }
    }
}
