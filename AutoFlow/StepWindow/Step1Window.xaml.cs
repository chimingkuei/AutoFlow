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
using System.Windows.Media.Media3D;
using System.Windows.Shapes;

namespace AutoFlow.StepWindow
{
    public class Parameter
    {
        public string VSM_Text_val { get; set; }
        public string DTCS_Text_val { get; set; }
    }

    public partial class Step1Window : Window
    {
        public Step1Window()
        {
            InitializeComponent();
        }

        #region Function
        #region Config
        private void LoadConfig()
        {
            List<Parameter> Parameter_info = Config.Load();
            VSM_Text.Text = Parameter_info[0].VSM_Text_val;
            DTCS_Text.Text = Parameter_info[0].DTCS_Text_val;
        }

        private void SaveConfig()
        {
            List<Parameter> Parameter_config = new List<Parameter>()
            {
                new Parameter() {
                    VSM_Text_val = VSM_Text.Text,
                    DTCS_Text_val = DTCS_Text.Text,
                }
            };
            Config.Save(Parameter_config);

        }
        #endregion
        #endregion

        #region Parameter and Init
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadConfig();
            Step1Data.SendValueEventHandler1 += (val) =>
            {
                VSM_Text.Text = val;
            };
            Step1Data.SendValueEventHandler2 += (val) =>
            {
                DTCS_Text.Text = val;
            };

        }
        BaseConfig<Parameter> Config = new BaseConfig<Parameter>(@"Step1Data.json");
        #endregion

        private void VSM_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool1 = true;
        }

        private void DTCS_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool2 = true;
        }

        private void Save_Config_Click(object sender, RoutedEventArgs e)
        {
            SaveConfig();
        }
    }
}
