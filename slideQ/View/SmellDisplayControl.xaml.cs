using slideQ.Model;
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
using slideQ;

namespace PowerPointAddInTest_CountSlides
{
    /// <summary>
    /// Interaction logic for SmellDisplayControl.xaml
    /// </summary>
    public partial class SmellDisplayControl : UserControl
    {
        public static ListView PPTSmellList;
        public SmellDisplayControl()
        {
            InitializeComponent();
            PPTSmellList = Smells;
        }

        private void Smells_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Smells_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var item = (sender as ListView).SelectedItem;
            if (item != null)
            {
                PresentationSmell data = item as PresentationSmell;
                Ribbon.Gotoslide(data.SlideNo);
            }
        }
    }
}
