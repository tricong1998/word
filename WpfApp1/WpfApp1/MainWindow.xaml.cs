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

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Words w = new Words();
            String path1 = txtClick.Text;
            String path2 = txtClick2.Text;
            //MessageBox.Show(w.compare(path1, path2));
            if (w.compare(path1, path2))
            {
                MessageBox.Show("Giong");
            }
            else
            {
                MessageBox.Show("K giong");
            }

        }

        private void btFont_Click(object sender, RoutedEventArgs e)
        {
            Words w = new Words();

        }
    }
}
