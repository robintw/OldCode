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
using System.Diagnostics;
using System.IO;

namespace TestWPF
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        private Results SearchResults;
        public Window1()
        {
            InitializeComponent();
        }

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            SearchResults.Add("HTCR.pdf", @"D:\Users\Robin Wilson\Documents\University\HTCR.pdf");
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SearchResults = new Results();
            listBox1.ItemsSource = SearchResults;
            comboBox1.ItemsSource = SearchResults;
            comboBox1.Focus();

            string[] files = Directory.GetFiles(@"C:\Users\Robin Wilson\AppData\Roaming\Microsoft\Windows\Start Menu", "*.lnk", SearchOption.AllDirectories);
            foreach (string filename in files)
            {
                FileInfo fi = new FileInfo(filename);

                SearchResults.Add(new Result(fi.Name, fi.FullName));
            }

            
        }

        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //if (e.AddedItems.Count > 0)
            //{
            //    Result SelRes = (Result)e.AddedItems[0];
            //    MessageBox.Show(SelRes.Name);
            //    //SelRes.Run();
            //}

        }

        private void comboBox1_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            comboBox1.IsDropDownOpen = true;
        }

        private void comboBox1_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                Result SelRes = (Result)comboBox1.SelectedItem;
                //MessageBox.Show(SelRes.Name);
                SelRes.Run();
            }
        }
    }
}
