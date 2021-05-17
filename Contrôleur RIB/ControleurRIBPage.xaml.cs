using System;
using System.Collections.Generic;
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
    /// Logique d'interaction pour ControleurRIBPage.xaml
    /// </summary>
    public partial class ControleurRIBPage : Page
    {
        public ControleurRIBPage()
        {
            InitializeComponent();
            DataContext = new VMControleurRIB();
        }
    }
}
