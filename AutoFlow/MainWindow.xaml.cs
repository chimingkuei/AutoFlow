using Newtonsoft.Json.Linq;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml;
using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static AutoFlow.BaseLogRecord;
using Color = System.Drawing.Color;
using OfficeOpenXml.Drawing;
using System.Drawing.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Interop;
using System.Diagnostics;
using AutoFlow.StepWindow;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Net.NetworkInformation;
using Path = System.IO.Path;

namespace AutoFlow
{
    public class Parameter
    {
        public List<System.Drawing.Point> click_points { get; set; }
        public string Setting_File_Location_val { get; set; }
        public string Ref_Fit_Location_val { get; set; }
        public string VSM_File_Location_val { get; set; }
        public string Xlsx_File_Location_val { get; set; }
        public string Wafer_Type_val { get; set; }
        public string VSM_Windows_Title_val { get; set; }
        public string VSM_Windows_X_val { get; set; }
        public string VSM_Windows_Y_val { get; set; }
        public string VSM_Windows_Width_val { get; set; }
        public string VSM_Windows_Height_val { get; set; }
        public string Dat_Windows_Title_val { get; set; }
        public string Dat_Windows_X_val { get; set; }
        public string Dat_Windows_Y_val { get; set; }
        public string Dat_Windows_Width_val { get; set; }
        public string Dat_Windows_Height_val { get; set; }
    }
    public partial class MainWindow : System.Windows.Window
    {
        
        public MainWindow()
        {
            InitializeComponent();
        }

        #region Function
        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (MessageBox.Show("請問是否要關閉？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

        #region NotifyIcon
        private System.Windows.Forms.NotifyIcon notifyIcon = null;
        System.Windows.Forms.ContextMenu nIconMenu = new System.Windows.Forms.ContextMenu();
        System.Windows.Forms.MenuItem nIconMenuItem1 = new System.Windows.Forms.MenuItem();
        System.Windows.Forms.MenuItem nIconMenuItem2 = new System.Windows.Forms.MenuItem();
        System.Windows.Forms.MenuItem nIconMenuItem3 = new System.Windows.Forms.MenuItem();
        System.Windows.Forms.MenuItem nIconMenuItem4 = new System.Windows.Forms.MenuItem();
        private void InitialTray()
        {
            notifyIcon = new System.Windows.Forms.NotifyIcon();
            notifyIcon.Icon = new System.Drawing.Icon(@"Icon/Deepwise.ico");
            notifyIcon.Text = "AutoFlow";
            notifyIcon.Visible = true;
            notifyIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(notifyIcon_MouseClick);
            this.StateChanged += new EventHandler(WPFUI_StateChanged);
            nIconMenuItem1.Index = 0;
            nIconMenuItem1.Text = "開始";
            nIconMenuItem1.Click += new System.EventHandler(nIconMenuItem1_Click);
            nIconMenu.MenuItems.Add(nIconMenuItem1);
            nIconMenuItem2.Index = 0;
            nIconMenuItem2.Text = "擷取螢幕";
            nIconMenuItem2.Click += new System.EventHandler(nIconMenuItem2_Click);
            nIconMenu.MenuItems.Add(nIconMenuItem2);
            nIconMenuItem3.Index = 0;
            nIconMenuItem3.Text = "自動點位抓取參數頁面";
            nIconMenuItem3.Click += new System.EventHandler(nIconMenuItem3_Click);
            nIconMenu.MenuItems.Add(nIconMenuItem3);
            nIconMenuItem4.Index = 0;
            nIconMenuItem4.Text = "晶圓點位參數頁面";
            nIconMenuItem4.Click += new System.EventHandler(nIconMenuItem4_Click);
            nIconMenu.MenuItems.Add(nIconMenuItem4);
            notifyIcon.ContextMenu = nIconMenu;
        }
        private void notifyIcon_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            //如果鼠标左键单击
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                if (this.Visibility == Visibility.Visible)
                {
                    this.Visibility = Visibility.Hidden;
                }
                else
                {
                    this.Show();
                    this.WindowState = (WindowState)System.Windows.Forms.FormWindowState.Normal;
                }

            }
        }
        private void WPFUI_StateChanged(object sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Minimized)
            {
                this.Visibility = Visibility.Hidden;
            }
        }
        private void nIconMenuItem1_Click(object sender, System.EventArgs e)
        {
            AutoStart();
        }
        private void nIconMenuItem2_Click(object sender, System.EventArgs e)
        {
            Do.CaptureScreen(Display_Image);
        }
        private void nIconMenuItem3_Click(object sender, System.EventArgs e)
        {
            OpenStep1Window();
        }
        private void nIconMenuItem4_Click(object sender, System.EventArgs e)
        {
            OpenWaferWindow();
        }
        #endregion

        #region Dispatcher Invoke Wrapper
        public string TextBoxDispatcherGetValue(TextBox control)
        {
            string name = "";
            this.Dispatcher.Invoke(() =>
            {
                name = control.Text;
            });
            return name;

        }

        public int IntegerUpDownDispatcherGetValue(Xceed.Wpf.Toolkit.IntegerUpDown control)
        {
            int value = new int();
            this.Dispatcher.Invoke(() =>
            {
                value = Convert.ToInt32(control.Text);
            });
            return value;
        }
        #endregion

        #region GetWindowRect
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion

        #region Config
        private void LoadConfig()
        {
            List<Parameter> Parameter_info = Config.Load();
            if (Parameter_info != null)
            {
                Setting_File_Location.Text = Parameter_info[0].Setting_File_Location_val;
                Ref_Fit_Location.Text = Parameter_info[0].Ref_Fit_Location_val;
                VSM_File_Location.Text = Parameter_info[0].VSM_File_Location_val;
                Xlsx_File_Location.Text = Parameter_info[0].Xlsx_File_Location_val;
                Wafer_Type.Text = Parameter_info[0].Wafer_Type_val;
                VSM_Windows_Title.Text = Parameter_info[0].VSM_Windows_Title_val;
                VSM_Windows_X.Text = Parameter_info[0].VSM_Windows_X_val;
                VSM_Windows_Y.Text = Parameter_info[0].VSM_Windows_Y_val;
                VSM_Windows_Width.Text = Parameter_info[0].VSM_Windows_Width_val;
                VSM_Windows_Height.Text = Parameter_info[0].VSM_Windows_Height_val;
                Dat_Windows_Title.Text = Parameter_info[0].Dat_Windows_Title_val;
                Dat_Windows_X.Text = Parameter_info[0].Dat_Windows_X_val;
                Dat_Windows_Y.Text = Parameter_info[0].Dat_Windows_Y_val;
                Dat_Windows_Width.Text = Parameter_info[0].Dat_Windows_Width_val;
                Dat_Windows_Height.Text = Parameter_info[0].Dat_Windows_Height_val;
            }
        }

        private void SaveConfig()
        {
            List<Parameter> Parameter_config = new List<Parameter>()
            {
                new Parameter() {
                    click_points = pointlist,
                    Setting_File_Location_val = Setting_File_Location.Text,
                    Ref_Fit_Location_val = Ref_Fit_Location.Text,
                    VSM_File_Location_val = VSM_File_Location.Text,
                    Xlsx_File_Location_val = Xlsx_File_Location.Text,
                    Wafer_Type_val = Wafer_Type.Text,
                    VSM_Windows_Title_val = VSM_Windows_Title.Text,
                    VSM_Windows_X_val = VSM_Windows_X.Text,
                    VSM_Windows_Y_val = VSM_Windows_Y.Text,
                    VSM_Windows_Width_val = VSM_Windows_Width.Text,
                    VSM_Windows_Height_val =VSM_Windows_Height.Text,
                    Dat_Windows_Title_val = Dat_Windows_Title.Text,
                    Dat_Windows_X_val = Dat_Windows_X.Text,
                    Dat_Windows_Y_val = Dat_Windows_Y.Text,
                    Dat_Windows_Width_val = Dat_Windows_Width.Text,
                    Dat_Windows_Height_val =Dat_Windows_Height.Text
                 }
            };
            Config.Save(Parameter_config);
        }

        private void LoadStep1WaferConfig(int group_num)
        {
            List<Parameter> Parameter_info = Config.Load();
            switch (Parameter_info[0].Wafer_Type_val)
            {
                case "6吋晶圓":
                    {
                        Step1Parameter_info = Step1Config.Load(group_num, 0);
                        WaferPointParameter_info = WaferPointConfig.Load(group_num, 0);
                        break;
                    }
                case "4吋晶圓":
                    {
                        Step1Parameter_info = Step1Config.Load(group_num, 1);
                        WaferPointParameter_info = WaferPointConfig.Load(group_num, 1);
                        break;
                    }
                case "3吋晶圓":
                    {
                        Step1Parameter_info = Step1Config.Load(group_num, 2);
                        WaferPointParameter_info = WaferPointConfig.Load(group_num, 2);
                        break;
                    }
            }
        }
        #endregion

        #region For Mouse Button Event function
        private void DrawCross()
        {
            Line cross1 = new Line
            {
                Stroke = System.Windows.Media.Brushes.Red,
                StrokeThickness = 3,
                X1 = _downPoint.X - 3,
                Y1 = _downPoint.Y - 3,
                X2 = _downPoint.X + 3,
                Y2 = _downPoint.Y + 3
            };
            Line cross2 = new Line
            {
                Stroke = System.Windows.Media.Brushes.Red,
                StrokeThickness = 3,
                X1 = _downPoint.X + 3,
                Y1 = _downPoint.Y - 3,
                X2 = _downPoint.X - 3,
                Y2 = _downPoint.Y + 3
            };
            cross1list.Add(cross1);
            cross2list.Add(cross2);
            Canvas.Children.Add(cross1);
            Canvas.Children.Add(cross2);
        }
        private void DrawDot()
        {
            Ellipse dot = new Ellipse
            {
                Stroke = System.Windows.Media.Brushes.Red,
                StrokeThickness = 5
            };
            Canvas.SetLeft(dot, _downPoint.X);
            Canvas.SetTop(dot, _downPoint.Y);
            dotlist.Add(dot);
            Canvas.Children.Add(dot);
        }
        private void GetClickPoint()
        {
            //DrawDot();
            DrawCross();
            // List螢幕座標位置
            int screen_x = Convert.ToInt32(_downPoint.X / Display_Image.ActualWidth * 1920);
            int screen_y = Convert.ToInt32(_downPoint.Y / Display_Image.ActualHeight * 1080);
            pointlist.Add(new System.Drawing.Point(screen_x, screen_y));
            Logger.WriteLog($"螢幕(X,Y)座標:({screen_x},{screen_y})", LogLevel.General, richTextBoxGeneral);
        }
        private void RemoveDot()
        {
            if (dotlist.Count != 0)
            {
                Canvas.Children.Remove(dotlist[dotlist.Count - 1]);
                dotlist.Remove(dotlist[dotlist.Count - 1]);
                pointlist.Remove(pointlist[pointlist.Count - 1]);
            }
        }
        private void RemoveCross()
        {
            if (cross1list.Count != 0)
            {
                Canvas.Children.Remove(cross1list[cross1list.Count - 1]);
                Canvas.Children.Remove(cross2list[cross2list.Count - 1]);
                cross1list.Remove(cross1list[cross1list.Count - 1]);
                cross2list.Remove(cross2list[cross2list.Count - 1]);
                pointlist.Remove(pointlist[pointlist.Count - 1]);
                if (pointlist.Count - 1 == -1)
                {
                    _downPoint.X = 0;
                    _downPoint.Y = 0;
                    Logger.WriteLog("螢幕(X,Y)座標:(0,0)", LogLevel.General, richTextBoxGeneral);
                }
                else
                {
                    int screen_x = pointlist[pointlist.Count - 1].X;
                    int screen_y = pointlist[pointlist.Count - 1].Y;
                    _downPoint.X = Convert.ToInt32(Convert.ToDouble(screen_x) / 1920 * Display_Image.ActualWidth);
                    _downPoint.Y = Convert.ToInt32(Convert.ToDouble(screen_y) / 1080 * Display_Image.ActualHeight);
                    Logger.WriteLog($"螢幕(X,Y)座標:({screen_x},{screen_y})", LogLevel.General, richTextBoxGeneral);
                }
               
            }
        }
        private void RemoveClickPoint()
        {
            //RemoveDot();
            RemoveCross();
        }
        #endregion

        #region Coordinate Format Conversion
        private string ConvertCoordStr(System.Windows.Point point , System.Windows.Controls.Image display_image)
        {
            if (point!=new System.Windows.Point(0,0))
            {
                string x = Convert.ToInt32(point.X / display_image.ActualWidth * 1920).ToString();
                string y = Convert.ToInt32(point.Y / display_image.ActualHeight * 1080).ToString();
                return "(" + x + "," + y + ")";
            }
            else
            {
                return "(0,0)";
            }
        }
        #endregion

        #region Action
        private static bool IsStep1WindowVisible(Step1Window window)
        {
            return window != null && window.IsVisible;
        }
        private void OpenStep1Window()
        {
            if (SWInstance == null || !IsStep1WindowVisible(SWInstance))
            {
                SWInstance = new Step1Window();
                SWInstance.Left = this.Left + (this.Width - SWInstance.Width) / 2;
                SWInstance.Top = this.Top + this.Height / 1.7;
                SWInstance.Show();
            }
            else
            {
                MessageBox.Show("已有一個自動點位抓取參數頁面視窗!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private static bool IsWaferWindowVisible(WaferWindow window)
        {
            return window != null && window.IsVisible;
        }
        private void OpenWaferWindow()
        {
            if (WWInstance == null || !IsWaferWindowVisible(WWInstance))
            {
                WWInstance = new WaferWindow();
                WWInstance.Left = this.Left + (this.Width - WWInstance.Width) / 2;
                WWInstance.Top = this.Top + this.Height / 1.7;
                WWInstance.Show();
            }
            else
            {
                MessageBox.Show("已有一個晶圓點位參數頁面視窗!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void AutoStart()
        {
            LoadStep1WaferConfig(3);
            string[] vsm_file = Do.GetFilename(TextBoxDispatcherGetValue(VSM_File_Location), "*.vsm");
            for (int file = 0; file < vsm_file.Length; file++)
            {
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].Open_Text_val, "點選資料夾圖示");
                Thread.Sleep(1000);
                Tuple<int, int, int, int> vsm_dialogue_windows_pos = new Tuple<int, int, int, int>(Convert.ToInt32(VSM_Windows_X.Text), Convert.ToInt32(VSM_Windows_Y.Text), Convert.ToInt32(VSM_Windows_Width.Text), Convert.ToInt32(VSM_Windows_Height.Text));
                Do.SetWindowsPosition(VSM_Windows_Title.Text, vsm_dialogue_windows_pos);
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].ChooseVSMPath_Text_val, "選擇vsm路徑");
                Do.SimulateInputText(TextBoxDispatcherGetValue(VSM_File_Location), "輸入vsm檔路徑");
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].VSM_Text_val, "點選開啟檔案類型欄位");
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].VSMType_Text_val, "選擇開檔類型vsm");
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].InputVSM_Text_val, "點選輸入vsm檔欄位");
                Do.SimulateInputText(Path.GetFileName(vsm_file[file]), "輸入vsm檔名");
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].TurnOn_Text_val, "點選開啟");
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].View_Text_val, "點選view");
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].Display_Text_val, "點選Display");
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].OnePane_Text_val, "點選1Pane");
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].Magnification_Text_val, "點選放大");
                Thread.Sleep(3000);
                foreach (var point in WaferPointParameter_info[0].WaferPoint_val)
                {
                    Do.ConvertWaferCoordStr(point);
                    Do.SimulateLeftMouseDoubleClick(Do.ConvertWaferCoordStr(point).Item2, "點選Wafer點位");
                    Do.SimulateLeftMouseClick(Step1Parameter_info[0].DTCS_Text_val, "點選DTCS");
                    Do.SimulateLeftMouseClick(Step1Parameter_info[0].OK_Text_val, "點選OK");
                    Do.SimulateLeftMouseClick(Step1Parameter_info[0].Save_Text_val, "點選Save");
                    Thread.Sleep(1000);
                    Tuple<int, int, int, int> dat_dialogue_windows_pos = new Tuple<int, int, int, int>(Convert.ToInt32(Dat_Windows_X.Text), Convert.ToInt32(Dat_Windows_Y.Text), Convert.ToInt32(Dat_Windows_Width.Text), Convert.ToInt32(Dat_Windows_Height.Text));
                    Do.SetWindowsPosition(Dat_Windows_Title.Text, dat_dialogue_windows_pos);
                    Do.SimulateLeftMouseClick(Step1Parameter_info[0].Dat_Text_val, "點選存檔類型欄位");
                    Do.SimulateLeftMouseClick(Step1Parameter_info[0].DatType_Text_val, "選擇儲存類型dat");
                    Do.SimulateLeftMouseClick(Step1Parameter_info[0].InputDat_Text_val, "點選輸入dat檔欄位");
                    Do.SimulateInputText(Path.GetFileNameWithoutExtension(vsm_file[file]) + Do.ConvertWaferCoordStr(point).Item1, "輸入dat檔名");
                    Do.SimulateLeftMouseClick(Step1Parameter_info[0].Archive_Text_val, "點選存檔");
                    Do.SimulateLeftMouseClick(Step1Parameter_info[0].CloseVDSW_Text_val, "關閉VDSW");
                    Thread.Sleep(100);
                }
                Do.SimulateLeftMouseClick(Step1Parameter_info[0].CloseWafer_Text_val, "關閉Wafer");
                Do.MoveFileToUpper(VSM_File_Location.Text, Path.Combine(Ref_Fit_Location.Text, "sample_spectrum"), "*dat");
                Do.RunSoftware(Ref_Fit_Location.Text);
                if (Do.CheckCSV(Path.Combine(Ref_Fit_Location.Text, "output_waveform.csv"), Path.Combine(Ref_Fit_Location.Text, "output_parameters.csv")))
                {
                    EH.waferID = Path.GetFileNameWithoutExtension(vsm_file[file]);
                    Do.DeleteFile(Path.Combine(Ref_Fit_Location.Text, "sample_spectrum"), "*dat");
                    EH.WaveToScatterChart(Path.Combine(Ref_Fit_Location.Text, "output_waveform.csv"), Path.Combine(Xlsx_File_Location.Text, Path.GetFileNameWithoutExtension(vsm_file[file]) + "output_waveform.xlsx"));
                    EH.ParameterToScatterChart(Path.Combine(Ref_Fit_Location.Text, "output_parameters.csv"), Path.Combine(Xlsx_File_Location.Text, Path.GetFileNameWithoutExtension(vsm_file[file]) + "output_parameters.xlsx"));
                }
                File.Move(Path.Combine(Ref_Fit_Location.Text, "output_waveform.csv"), Path.Combine(Xlsx_File_Location.Text, Path.GetFileNameWithoutExtension(vsm_file[file]) + "_output_waveform.csv"));
                File.Move(Path.Combine(Ref_Fit_Location.Text, "output_parameters.csv"), Path.Combine(Xlsx_File_Location.Text, Path.GetFileNameWithoutExtension(vsm_file[file]) + "_output_parameters.csv"));
            };
        }
        #endregion

        private void CheckSendValueInit()
        {
            Step1Data.CheckSendValueEventHandler1 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data1 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler2 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data2 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler3 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data3 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler4 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data4 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler5 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data5 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler6 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data6 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler7 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data7 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler8 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data8 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler9 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data9 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler10 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data10 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler11 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data11 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler12 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data12 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler13 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data13 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler14 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data14 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler15 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data15 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler16 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data16 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler17 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data17 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler18 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data18 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler19 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data19 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler20 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data20 = ConvertCoordStr(_downPoint, Display_Image);
                }
            };
        }
        #endregion

        #region Parameter and Init
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            InitialTray();
            LoadConfig();
            //Do.LoadEIM();
            //Do.CloseCapsLock();
            CheckSendValueInit();
            Model_Type_state = true;
        }
        BaseConfig<Parameter> Config = new BaseConfig<Parameter>(@"Config\Config.json");
        BaseConfig<AutoFlow.StepWindow.Step1Parameter> Step1Config = new BaseConfig<AutoFlow.StepWindow.Step1Parameter>(@"Config\Step1Config.json");
        BaseConfig<AutoFlow.StepWindow.WaferPointParameter> WaferPointConfig = new BaseConfig<AutoFlow.StepWindow.WaferPointParameter>(@"Config\WaferPoint.json");
        Core Do = new Core();
        ExcelHandler EH = new ExcelHandler();
        BaseLogRecord Logger = new BaseLogRecord();
        private bool _started;
        private System.Windows.Point _downPoint;
        private static Step1Window SWInstance;
        private static WaferWindow WWInstance;
        private bool Model_Type_state = false;
        List<AutoFlow.StepWindow.Step1Parameter> Step1Parameter_info;
        List<AutoFlow.StepWindow.WaferPointParameter> WaferPointParameter_info;
        #endregion

        #region Main Screen
        private void Main_Btn_Click(object sender, RoutedEventArgs e)
        {
            switch ((sender as Button).Name)
            {
                case nameof(Start):
                    {
                        AutoStart();
                        break;
                    }
                case nameof(Capture_Screen):
                    {
                        Do.CaptureScreen(Display_Image);
                        break;
                    }
                case nameof(Open_Step1Window):
                    {
                        OpenStep1Window();
                        break;
                    }
                case nameof(Open_Wafer_Point):
                    {
                        OpenWaferWindow();
                        break;
                    }
            }
        }
        #endregion

        #region Parameter Screen
        private void Parameter_Btn_Click(object sender, RoutedEventArgs e)
        {
            switch ((sender as Button).Name)
            {
                case nameof(Open_Setting_File_Path):
                    {
                        OpenFileDialog open_setting_file_path = new OpenFileDialog();
                        open_setting_file_path.Title = "選擇檔案";
                        open_setting_file_path.Filter = "文本檔案 (*.json)|*.json|所有檔案 (*.*)|*.*";
                        if (open_setting_file_path.ShowDialog()==true)
                        {
                            Setting_File_Location.Text = open_setting_file_path.FileName;
                        }
                        break;
                    }
                case nameof(Open_VSM_Folder):
                    {
                        System.Windows.Forms.FolderBrowserDialog open_vsm_folder_path = new System.Windows.Forms.FolderBrowserDialog();
                        open_vsm_folder_path.Description = "選擇vsm檔資料夾";
                        open_vsm_folder_path.ShowDialog();
                        Setting_File_Location.Text = open_vsm_folder_path.SelectedPath;
                        break;
                    }
                case nameof(Open_Ref_Fit_Folder):
                    {
                        System.Windows.Forms.FolderBrowserDialog open_vsm_folder_path = new System.Windows.Forms.FolderBrowserDialog();
                        open_vsm_folder_path.Description = "選擇Ref-Fit軟體資料夾";
                        open_vsm_folder_path.ShowDialog();
                        Ref_Fit_Location.Text = open_vsm_folder_path.SelectedPath;
                        break;
                    }
                case nameof(Open_Xlsx_Folder):
                    {
                        System.Windows.Forms.FolderBrowserDialog open_xlsx_folder_path = new System.Windows.Forms.FolderBrowserDialog();
                        open_xlsx_folder_path.Description = "選擇xlsx檔資料夾";
                        open_xlsx_folder_path.ShowDialog();
                        Xlsx_File_Location.Text = open_xlsx_folder_path.SelectedPath;
                        break;
                    }
                case nameof(VSM_Set_Windows):
                    {
                        int x = Convert.ToInt32(VSM_Windows_X.Text);
                        int y = Convert.ToInt32(VSM_Windows_Y.Text);
                        int w = Convert.ToInt32(VSM_Windows_Width.Text);
                        int h = Convert.ToInt32(VSM_Windows_Height.Text);
                        Do.SetWindowsPosition(VSM_Windows_Title.Text, new Tuple<int, int, int, int>(x, y, w, h));
                        break;
                    }
                case nameof(Dat_Set_Windows):
                    {
                        int x = Convert.ToInt32(Dat_Windows_X.Text);
                        int y = Convert.ToInt32(Dat_Windows_Y.Text);
                        int w = Convert.ToInt32(Dat_Windows_Width.Text);
                        int h = Convert.ToInt32(Dat_Windows_Height.Text);
                        Do.SetWindowsPosition(Dat_Windows_Title.Text, new Tuple<int, int, int, int>(x, y, w, h));
                        break;
                    }
                case nameof(Save_Config):
                    {
                        SaveConfig();
                        Logger.WriteLog("儲存參數!", LogLevel.General, richTextBoxGeneral);
                        break;
                    }
            }
        }
        #endregion

        #region Mouse Button Event
        private List<Ellipse> dotlist = new List<Ellipse>();
        private List<Line> cross1list = new List<Line>();
        private List<Line> cross2list = new List<Line>();
        private List<System.Drawing.Point> pointlist = new List<System.Drawing.Point>();
        private void AddPointButtonDown(object sender, MouseButtonEventArgs e)
        {
            _started = true;
            _downPoint = e.GetPosition(Display_Image);
            GetClickPoint();
        }
        private void AddPointButtonUp(object sender, MouseButtonEventArgs e)
        {
            _started = false;
        }
        private void DeletePointButtonDown(object sender, MouseButtonEventArgs e)
        {
            _started = true;
            RemoveClickPoint();
        }
        private void DeletePointButtonUp(object sender, MouseButtonEventArgs e)
        {
            _started = false;
        }
        #endregion

        #region Combobox SelectionChanged
        private void Wafer_Type_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Model_Type_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Model_Type_state)
            {
                if (!string.IsNullOrEmpty(Setting_File_Location.Text))
                {
                    if (File.Exists(Setting_File_Location.Text))
                    {
                        string stringToRemove = "System.Windows.Controls.ComboBoxItem: ";
                        Do.CheckModel(Setting_File_Location.Text, Model_Type.SelectedValue.ToString().Replace(stringToRemove, ""));
                    }
                    else
                    {
                        MessageBox.Show("setting檔案不存在!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("請輸入setting檔案位置!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                }


            }
        }
        #endregion

    }
}
