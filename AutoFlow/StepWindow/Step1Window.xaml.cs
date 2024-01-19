using AutoFlow.StepWindow;
using System;
using System.Collections.Generic;
using System.Drawing;
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
        public string Open_Text_val { get; set; }
        public string ChoosePath_Text_val { get; set; }
        public string VSM_Text_val { get; set; }
        public string VSMType_Text_val { get; set; }
        public string InputVSM_Text_val { get; set; }
        public string TurnOn_Text_val { get; set; }
        public string View_Text_val { get; set; }
        public string Display_Text_val { get; set; }
        public string OnePane_Text_val { get; set; }
        public string Magnification_Text_val { get; set; }
        public string DTCS_Text_val { get; set; }
        public string OK_Text_val { get; set; }
        public string Save_Text_val { get; set; }
        public string Dat_Text_val { get; set; }
        public string DatType_Text_val { get; set; }
        public string InputDat_Text_val { get; set; }
        public string Archive_Text_val { get; set; }
        public string CloseVDSW_Text_val { get; set; }
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
            Open_Text.Text = Parameter_info[0].Open_Text_val;
            ChoosePath_Text.Text = Parameter_info[0].ChoosePath_Text_val;
            VSM_Text.Text = Parameter_info[0].VSM_Text_val;
            VSMType_Text.Text = Parameter_info[0].VSMType_Text_val;
            InputVSM_Text.Text = Parameter_info[0].InputVSM_Text_val;
            TurnOn_Text.Text = Parameter_info[0].TurnOn_Text_val;
            View_Text.Text = Parameter_info[0].View_Text_val;
            Display_Text.Text = Parameter_info[0].Display_Text_val;
            OnePane_Text.Text = Parameter_info[0].OnePane_Text_val;
            Magnification_Text.Text = Parameter_info[0].Magnification_Text_val;
            DTCS_Text.Text = Parameter_info[0].DTCS_Text_val;
            OK_Text.Text = Parameter_info[0].OK_Text_val;
            Save_Text.Text = Parameter_info[0].Save_Text_val;
            Dat_Text.Text = Parameter_info[0].Dat_Text_val;
            DatType_Text.Text = Parameter_info[0].DatType_Text_val;
            InputDat_Text.Text = Parameter_info[0].InputDat_Text_val;
            Archive_Text.Text = Parameter_info[0].Archive_Text_val;
            CloseVDSW_Text.Text =  Parameter_info[0].CloseVDSW_Text_val;
        }

        private void SaveConfig()
        {
            List<Parameter> Parameter_config = new List<Parameter>()
            {
                new Parameter() {
                    Open_Text_val=Open_Text.Text,
                    ChoosePath_Text_val=ChoosePath_Text.Text,
                    VSM_Text_val=VSM_Text.Text,
                    VSMType_Text_val=VSMType_Text.Text,
                    InputVSM_Text_val=InputVSM_Text.Text,
                    TurnOn_Text_val=TurnOn_Text.Text,
                    View_Text_val=View_Text.Text,
                    Display_Text_val=Display_Text.Text,
                    OnePane_Text_val=OnePane_Text.Text,
                    Magnification_Text_val=Magnification_Text.Text,
                    DTCS_Text_val=DTCS_Text.Text,
                    OK_Text_val=OK_Text.Text,
                    Save_Text_val=Save_Text.Text,
                    Dat_Text_val=Dat_Text.Text,
                    DatType_Text_val=DatType_Text.Text,
                    InputDat_Text_val=InputDat_Text.Text,
                    Archive_Text_val=Archive_Text.Text,
                    CloseVDSW_Text_val=CloseVDSW_Text.Text
                   }
            };
            Config.Save(Parameter_config);
        }
        #endregion

        private void CheckSendValueInit()
        {
            Step1Data.SendValueEventHandler1 += (val) =>
            {
                Open_Text.Text = val;
            };
            Step1Data.SendValueEventHandler2 += (val) =>
            {
                ChoosePath_Text.Text = val;
            };
            Step1Data.SendValueEventHandler3 += (val) =>
            {
                VSM_Text.Text = val;
            };
            Step1Data.SendValueEventHandler4 += (val) =>
            {
                VSMType_Text.Text = val;
            };
            Step1Data.SendValueEventHandler5 += (val) =>
            {
                InputVSM_Text.Text = val;
            };
            Step1Data.SendValueEventHandler6 += (val) =>
            {
                TurnOn_Text.Text = val;
            };
            Step1Data.SendValueEventHandler7 += (val) =>
            {
                View_Text.Text = val;
            };
            Step1Data.SendValueEventHandler8 += (val) =>
            {
                Display_Text.Text = val;
            };
            Step1Data.SendValueEventHandler9 += (val) =>
            {
                OnePane_Text.Text = val;
            };
            Step1Data.SendValueEventHandler10 += (val) =>
            {
                Magnification_Text.Text = val;
            };
            Step1Data.SendValueEventHandler11 += (val) =>
            {
                DTCS_Text.Text = val;
            };
            Step1Data.SendValueEventHandler12 += (val) =>
            {
                OK_Text.Text = val;
            };
            Step1Data.SendValueEventHandler13 += (val) =>
            {
                Save_Text.Text = val;
            };
            Step1Data.SendValueEventHandler14 += (val) =>
            {
                Dat_Text.Text = val;
            };
            Step1Data.SendValueEventHandler15 += (val) =>
            {
                DatType_Text.Text = val;
            };
            Step1Data.SendValueEventHandler16 += (val) =>
            {
                InputDat_Text.Text = val;
            };
            Step1Data.SendValueEventHandler17 += (val) =>
            {
                Archive_Text.Text = val;
            };
            Step1Data.SendValueEventHandler18 += (val) =>
            {
                CloseVDSW_Text.Text = val;
            };
        }
        #endregion

        #region Parameter and Init
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadConfig();
            CheckSendValueInit();
        }
        BaseConfig<Parameter> Config = new BaseConfig<Parameter>(@"Step1Data.json");
        #endregion

        private void Save_Config_Click(object sender, RoutedEventArgs e)
        {
            SaveConfig();
        }

        private void Open_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool1 = true;
        }

        private void ChoosePath_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool2 = true;
        }

        private void VSM_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool3 = true;
        }

        private void VSMType_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool4 = true;
        }

        private void InputVSM_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool5 = true;
        }

        private void TurnOn_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool6 = true;
        }

        private void View_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool7 = true;
        }

        private void Display_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool8 = true;
        }

        private void OnePane_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool9 = true;
        }

        private void Magnification_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool10 = true;
        }

        private void DTCS_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool11 = true;
        }

        private void OK_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool12 = true;
        }

        private void Save_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool13 = true;
        }
        private void Dat_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool14 = true;
        }
        private void DatType_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool15 = true;
        }

        private void InputDat_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool16 = true;
        }

        private void Archive_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool17 = true;
        }

        private void CloseVDSW_Checked(object sender, RoutedEventArgs e)
        {
            Step1Data.Step1_bool18 = true;
        }

       
    }
}
