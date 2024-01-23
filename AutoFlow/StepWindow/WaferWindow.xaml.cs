using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    public class WaferParameter
    {
        public string WaferPoint_Csv_Path_val { get; set; }
        public string CoordX_val { get; set; }
        public string CoordY_val { get; set; }
        public string Origin_val { get; set; }
    }
    public class WaferPointParameter
    {
        public List<string> WaferPoint_val { get; set; }
    }
    public partial class WaferWindow : Window
    {
        public WaferWindow()
        {
            InitializeComponent();
        }

        #region Function
        #region Config
        private void SaveWaferPointConfig()
        {
            List<WaferPointParameter> WaferPointParameter_config = new List<WaferPointParameter>()
            {
                new WaferPointParameter() {
                                          WaferPoint_val = Temp
                }
            };
            WaferPoint.Save(WaferPointParameter_config);
        }

        private void LoadWaferConfig()
        {
            List<WaferParameter> WaferParameter_info = Wafer.Load();
            WaferPoint_Csv_Path.Text = WaferParameter_info[0].WaferPoint_Csv_Path_val;
            CoordX.Text = WaferParameter_info[0].CoordX_val;
            CoordY.Text = WaferParameter_info[0].CoordY_val;
            Origin.Text = WaferParameter_info[0].Origin_val;
        }

        private void SaveWaferConfig(int index)
        {
            List<WaferParameter> WaferParameter_config = new List<WaferParameter>()
            {
                new WaferParameter() {
                                          WaferPoint_Csv_Path_val = WaferPoint_Csv_Path.Text,
                                          CoordX_val = CoordX.Text,
                                          CoordY_val = CoordY.Text,
                                          Origin_val = Origin.Text
                }
            };
            Wafer.Save(WaferParameter_config, index);
        }
        #endregion

        private System.Drawing.Point ConvertCoordXY(string coord_str)
        {
            Match match = Regex.Match(coord_str, @"\((\d+),(\d+)\)");
            return new System.Drawing.Point(int.Parse(match.Groups[1].Value), int.Parse(match.Groups[2].Value));
        }
        #endregion

        #region Parameter and Init
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //LoadWaferConfig();
        }
        Core Do = new Core();
        ExcelHandler EH = new ExcelHandler();
        BaseConfig<WaferPointParameter> WaferPoint = new BaseConfig<WaferPointParameter>(@"Config\WaferPoint.json");
        BaseConfig<WaferParameter> Wafer = new BaseConfig<WaferParameter>(@"Config\WaferConfig.json");
        private List<string> Temp;
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
                        if (!string.IsNullOrEmpty(CoordX.Text))
                        {
                            if (!string.IsNullOrEmpty(CoordX.Text))
                            {
                                int x = Convert.ToInt32(CoordX.Text);
                                int y = Convert.ToInt32(CoordY.Text);
                                Do.SimulateLeftMouseClick(new System.Drawing.Point(x, y));
                            }
                            else
                            {
                                System.Windows.MessageBox.Show("請輸入Y座標!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }
                        else
                        {
                            System.Windows.MessageBox.Show("請輸入X座標!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        break;
                    }
                case nameof(Convert_Screen_Coordinate):
                    {
                        if (!string.IsNullOrEmpty(Origin.Text))
                        {
                            System.Drawing.Point origin = ConvertCoordXY(Origin.Text);
                            EH.ConvertScreenCoordinate(WaferPoint_Csv_Path.Text, new Tuple<int, int>(origin.X, origin.Y));
                        }
                        else
                        {
                            System.Windows.MessageBox.Show("請輸入原點座標!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        break;
                    }
                case nameof(Convert_WaferPoint_Json):
                    {
                        Temp = EH.ReadCsv(WaferPoint_Csv_Path.Text, EH.ConvertWaferPointJsonFormat);
                        SaveWaferPointConfig();
                        break;
                    }
                case nameof(Save_Config):
                    {
                        SaveWaferConfig(0);
                        break;
                    }
            }
        }
        #endregion

    }
}
