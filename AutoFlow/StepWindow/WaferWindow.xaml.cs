using Microsoft.Win32;
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
using System.Windows.Media.Media3D;
using System.Windows.Shapes;

namespace AutoFlow.StepWindow
{
    public class WaferPointParameter
    {
        public List<string> WaferPoint_val { get; set; }
        public string WaferPoint_Csv_Path_val { get; set; }
    }
    public partial class WaferWindow : Window
    {
        public WaferWindow()
        {
            InitializeComponent();
        }

        #region Function
        private void LoadConfig()
        {
            List<WaferPointParameter> WaferPointParameter_info = WaferPoint.Load();
            WaferPoint_Csv_Path.Text = WaferPointParameter_info[0].WaferPoint_Csv_Path_val;
        }

        private void SaveConfig()
        {
            List<WaferPointParameter> WaferPointParameter_config = new List<WaferPointParameter>()
            {
                new WaferPointParameter() {
                                          WaferPoint_Csv_Path_val = WaferPoint_Csv_Path.Text,
                                          WaferPoint_val = Temp
                }
            };
            WaferPoint.Save(WaferPointParameter_config);
        }
        #endregion

        #region Parameter and Init
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadConfig();
        }
        Core Do = new Core();
        ExcelHandler EH = new ExcelHandler();
        BaseConfig<WaferPointParameter> WaferPoint = new BaseConfig<WaferPointParameter>(@"WaferPointTest.json");
        private List<string> Temp { get; set; } = null;
        #endregion

        #region WaferWindow Screen
        private void Main_Btn_Click(object sender, RoutedEventArgs e)
        {
            switch ((sender as Button).Name)
            {
                case nameof(Open_WaferPoint_Csv_Path):
                    {
                        OpenFileDialog open_waferpoint_csv_path = new OpenFileDialog();
                        open_waferpoint_csv_path.Title = "選擇檔案";
                        open_waferpoint_csv_path.Filter = "文本檔案 (*.csv)|*.csv|所有檔案 (*.*)|*.*";
                        if (open_waferpoint_csv_path.ShowDialog() == true)
                        {
                            WaferPoint_Csv_Path.Text = open_waferpoint_csv_path.FileName;
                        }
                        break;
                    }
                case nameof(Move_Mouse):
                    {
                        int x = Convert.ToInt32(CoordX.Text);
                        int y = Convert.ToInt32(CoordY.Text);
                        Do.SimulateLeftMouseClick(new System.Drawing.Point(x, y));
                        break;
                    }
                case nameof(Csv_To_WaferPoint_Json):
                    {
                        Temp = EH.CSVToWaferPointJson(@"D:\Chimingkuei\repos\Project\AutoFlow\AutoFlow\bin\x64\Debug\Wafer click coordinate position.csv");
                        SaveConfig();
                        break;
                    }

            }
        }
        #endregion

    }
}
