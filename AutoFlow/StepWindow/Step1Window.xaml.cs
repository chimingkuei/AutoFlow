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
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AutoFlow.StepWindow
{
    /// <summary>
    /// Step1Window.xaml 的互動邏輯
    /// </summary>
    public partial class Step1Window : Window
    {
        public Step1Window()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Step1Data.SendValueEventHandler1 += (val) =>
            {
                textBox1.Text = val;
            };
            Step1Data.SendValueEventHandler2 += (val) =>
            {
                textBox2.Text = val;
            };

        }

        private void radioButton_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool1 = true;
        }
    }
}
