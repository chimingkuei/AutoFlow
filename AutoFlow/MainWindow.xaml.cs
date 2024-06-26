﻿using Newtonsoft.Json.Linq;
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
using Newtonsoft.Json;
using System.Management;
using System.Runtime.InteropServices.ComTypes;

namespace AutoFlow
{
    public class Parameter
    {
        public string Setting_File_Location_val { get; set; }
        public string Ref_Fit_Location_val { get; set; }
        public string VSM_File_Location_val { get; set; }
        public string Xlsx_File_Location_val { get; set; }
        public string Wafer_Type_val { get; set; }
        public string Model_Type_val { get; set; }
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
        public string VDSW_Windows_Title_val { get; set; }
        public string VDSW_Windows_X_val { get; set; }
        public string VDSW_Windows_Y_val { get; set; }
        public string VDSW_Windows_Width_val { get; set; }
        public string VDSW_Windows_Height_val { get; set; }
        public bool Save_Datfile_val { get; set; }
        public bool Fixed_Time_val { get; set; }
        public bool Stop_Writing_val { get; set; }
        public string Fixed_Time_Value_val { get; set; }
        public string Stop_Writing_Value_val { get; set; }
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

        public bool CheckBoxDispatcherGetValue(CheckBox control)
        {
            bool check = new bool();
            this.Dispatcher.Invoke(() =>
            {
                check = (bool)control.IsChecked;
            });
            return check;

        }

        public bool RadioButtonDispatcherGetValue(RadioButton control)
        {
            bool check = new bool();
            this.Dispatcher.Invoke(() =>
            {
                check = (bool)control.IsChecked;
            });
            return check;

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

        #region Software Lock
        private string GetBoardSerialNumber()
        {
            string boardSerial = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");
            foreach (ManagementObject mo in searcher.Get())
            {
                boardSerial = mo["SerialNumber"].ToString();
                break;
            }
            return boardSerial;
        }

        private void Lock()
        {
            if (GetBoardSerialNumber() != "07D5910_M71E680889")
            {
                MessageBox.Show("請聯繫廠商提供Licence!", "確認", MessageBoxButton.OK, MessageBoxImage.Warning);
                Environment.Exit(0);
            }
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
                Model_Type.Text = Parameter_info[0].Model_Type_val;
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
                VDSW_Windows_Title.Text = Parameter_info[0].VDSW_Windows_Title_val;
                VDSW_Windows_X.Text = Parameter_info[0].VDSW_Windows_X_val;
                VDSW_Windows_Y.Text = Parameter_info[0].VDSW_Windows_Y_val;
                VDSW_Windows_Width.Text = Parameter_info[0].VDSW_Windows_Width_val;
                VDSW_Windows_Height.Text = Parameter_info[0].VDSW_Windows_Height_val;
                Save_Datfile.IsChecked = Parameter_info[0].Save_Datfile_val;
                Fixed_Time.IsChecked = Parameter_info[0].Fixed_Time_val;
                Stop_Writing.IsChecked = Parameter_info[0].Stop_Writing_val;
                Fixed_Time_Value.Text = Parameter_info[0].Fixed_Time_Value_val;
                Stop_Writing_Value.Text = Parameter_info[0].Stop_Writing_Value_val;
            }
        }

        private void SaveConfig()
        {
            List<Parameter> Parameter_config = new List<Parameter>()
            {
                new Parameter() {
                    Setting_File_Location_val = Setting_File_Location.Text,
                    Ref_Fit_Location_val = Ref_Fit_Location.Text,
                    VSM_File_Location_val = VSM_File_Location.Text,
                    Xlsx_File_Location_val = Xlsx_File_Location.Text,
                    Wafer_Type_val = Wafer_Type.Text,
                    Model_Type_val = Model_Type.Text,
                    VSM_Windows_Title_val = VSM_Windows_Title.Text,
                    VSM_Windows_X_val = VSM_Windows_X.Text,
                    VSM_Windows_Y_val = VSM_Windows_Y.Text,
                    VSM_Windows_Width_val = VSM_Windows_Width.Text,
                    VSM_Windows_Height_val =VSM_Windows_Height.Text,
                    Dat_Windows_Title_val = Dat_Windows_Title.Text,
                    Dat_Windows_X_val = Dat_Windows_X.Text,
                    Dat_Windows_Y_val = Dat_Windows_Y.Text,
                    Dat_Windows_Width_val = Dat_Windows_Width.Text,
                    Dat_Windows_Height_val = Dat_Windows_Height.Text,
                    VDSW_Windows_Title_val = VDSW_Windows_Title.Text,
                    VDSW_Windows_X_val = VDSW_Windows_X.Text,
                    VDSW_Windows_Y_val = VDSW_Windows_Y.Text,
                    VDSW_Windows_Width_val = VDSW_Windows_Width.Text,
                    VDSW_Windows_Height_val = VDSW_Windows_Height.Text,
                    Save_Datfile_val = (bool)Save_Datfile.IsChecked,
                    Fixed_Time_val = (bool)Fixed_Time.IsChecked,
                    Stop_Writing_val = (bool)Stop_Writing.IsChecked,
                    Fixed_Time_Value_val = Fixed_Time_Value.Text,
                    Stop_Writing_Value_val = Stop_Writing_Value.Text
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
            Screenpointlist.Add(new System.Drawing.Point(screen_x, screen_y));
            pointlist.Add(_downPoint);
            Logger.WriteLog($"螢幕(X,Y)座標:({screen_x},{screen_y})", LogLevel.General, richTextBoxGeneral);
        }
        private void RemoveDot()
        {
            if (dotlist.Count != 0)
            {
                Canvas.Children.Remove(dotlist[dotlist.Count - 1]);
                dotlist.Remove(dotlist[dotlist.Count - 1]);
                Screenpointlist.Remove(Screenpointlist[Screenpointlist.Count - 1]);
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
                Screenpointlist.Remove(Screenpointlist[Screenpointlist.Count - 1]);
                pointlist.Remove(pointlist[pointlist.Count - 1]);
                if (Screenpointlist.Count == 0)
                {
                    _downPoint.X = 0;
                    _downPoint.Y = 0;
                    Logger.WriteLog("螢幕(X,Y)座標:(0,0)", LogLevel.General, richTextBoxGeneral);
                }
                else
                {
                    int screen_x = Screenpointlist[Screenpointlist.Count - 1].X;
                    int screen_y = Screenpointlist[Screenpointlist.Count - 1].Y;
                    _downPoint = pointlist[pointlist.Count - 1];
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
                Logger.WriteLog("開啟自動點位抓取頁面!", LogLevel.General, richTextBoxGeneral);
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
                Logger.WriteLog("開啟晶圓點位頁面!", LogLevel.General, richTextBoxGeneral);
            }
            else
            {
                MessageBox.Show("已有一個晶圓點位參數頁面視窗!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void SetWindowsPos(TextBox Windows_Title, Xceed.Wpf.Toolkit.IntegerUpDown Windows_X, Xceed.Wpf.Toolkit.IntegerUpDown Windows_Y, Xceed.Wpf.Toolkit.IntegerUpDown Windows_Width, Xceed.Wpf.Toolkit.IntegerUpDown Windows_Height, string annotation = null)
        {
            Thread.Sleep(500);
            int x = IntegerUpDownDispatcherGetValue(Windows_X);
            int y = IntegerUpDownDispatcherGetValue(Windows_Y);
            int w = IntegerUpDownDispatcherGetValue(Windows_Width);
            int h = IntegerUpDownDispatcherGetValue(Windows_Height);
            Tuple<int, int, int, int> vsm_dialogue_windows_pos = new Tuple<int, int, int, int>(x, y, w, h);
            Do.SetWindowsPosition(TextBoxDispatcherGetValue(Windows_Title), vsm_dialogue_windows_pos);
        }

        private void SetWindowsPosVDSW(string wafername, Xceed.Wpf.Toolkit.IntegerUpDown Windows_X, Xceed.Wpf.Toolkit.IntegerUpDown Windows_Y, Xceed.Wpf.Toolkit.IntegerUpDown Windows_Width, Xceed.Wpf.Toolkit.IntegerUpDown Windows_Height, string annotation = null)
        {
            Thread.Sleep(500);
            Thread.Sleep(1000);
            int x = IntegerUpDownDispatcherGetValue(Windows_X);
            int y = IntegerUpDownDispatcherGetValue(Windows_Y);
            int w = IntegerUpDownDispatcherGetValue(Windows_Width);
            int h = IntegerUpDownDispatcherGetValue(Windows_Height);
            Tuple<int, int, int, int> vsm_dialogue_windows_pos = new Tuple<int, int, int, int>(x, y, w, h);
            Do.SetWindowsPosition("VCSEL/DBR Spectrum - [" + Path.GetFileName(wafername) + "]", vsm_dialogue_windows_pos);
        }

        private string CreateDateDir()
        {
            string date = DateTime.Now.ToString("yyyyMMddhhmmss").ToString();
            string dir = Path.Combine(TextBoxDispatcherGetValue(Xlsx_File_Location), date);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            return date;
        }

        private Dictionary<string, string> SetCsvWorkPath(string vsm_file, string date)
        {
           
            Dictionary<string, string> dict = new Dictionary<string, string>();
            dict.Add("WaveCsvPath", Path.Combine(TextBoxDispatcherGetValue(Ref_Fit_Location), "output_waveform.csv"));
            dict.Add("ParameterCsvPath", Path.Combine(TextBoxDispatcherGetValue(Ref_Fit_Location), "output_parameters.csv"));
            dict.Add("sample_spectrum", Path.Combine(TextBoxDispatcherGetValue(Ref_Fit_Location), "sample_spectrum"));
            dict.Add("MoveWaveCsvPath", Path.Combine(TextBoxDispatcherGetValue(Xlsx_File_Location), date, Path.GetFileNameWithoutExtension(vsm_file) + "_output_waveform.csv"));
            dict.Add("MoveParameterCsvPath", Path.Combine(TextBoxDispatcherGetValue(Xlsx_File_Location), date, Path.GetFileNameWithoutExtension(vsm_file) + "_output_parameters.csv"));
            dict.Add("WaveXlsxPath", Path.Combine(TextBoxDispatcherGetValue(Xlsx_File_Location), date, Path.GetFileNameWithoutExtension(vsm_file) + "_output_waveform.xlsx"));
            dict.Add("ParameterXlsxPath", Path.Combine(TextBoxDispatcherGetValue(Xlsx_File_Location), date, Path.GetFileNameWithoutExtension(vsm_file) + "_output_parameters.xlsx"));
            dict.Add("DatePath", Path.Combine(TextBoxDispatcherGetValue(Xlsx_File_Location), date));
            return dict;
        }

        private void RemoveExtraData()
        {
            Do.DeleteFile(Path.Combine(TextBoxDispatcherGetValue(Ref_Fit_Location), "sample_spectrum"), "*dat");
            Do.DeleteFile(TextBoxDispatcherGetValue(VSM_File_Location), "*dat");
            string waveform_csv = Path.Combine(TextBoxDispatcherGetValue(Ref_Fit_Location), "output_waveform.csv");
            string parameter_csv = Path.Combine(TextBoxDispatcherGetValue(Ref_Fit_Location), "output_parameters.csv");
            if (File.Exists(waveform_csv))
                File.Delete(waveform_csv);
            if (File.Exists(parameter_csv))
                File.Delete(parameter_csv);
        }

        private void AutoStart()
        {
            cts = new CancellationTokenSource();
            Task.Run(() =>
            {
                #region Check Work Folder Path
                if (string.IsNullOrEmpty(TextBoxDispatcherGetValue(Setting_File_Location)))
                {
                    MessageBox.Show("請輸入setting檔案路徑!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (string.IsNullOrEmpty(TextBoxDispatcherGetValue(Ref_Fit_Location)))
                {
                    MessageBox.Show("請輸入Ref-Fit軟體資料夾路徑!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (string.IsNullOrEmpty(TextBoxDispatcherGetValue(VSM_File_Location)))
                {
                    MessageBox.Show("請輸入vsm檔資料夾路徑!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (string.IsNullOrEmpty(TextBoxDispatcherGetValue(Xlsx_File_Location)))
                {
                    MessageBox.Show("請輸入xlsx檔資料夾路徑!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                #endregion
                RemoveExtraData();
                LoadStep1WaferConfig(3);
                string[] vsm_file = Do.GetFilename(TextBoxDispatcherGetValue(VSM_File_Location), "*.vsm");
                if (vsm_file.Length != 0)
                {
                    string date = CreateDateDir();
                    string method = RadioButtonDispatcherGetValue(Fixed_Time)? "FixedTime": "StopWriting";
                    int timeout = RadioButtonDispatcherGetValue(Fixed_Time) ? Convert.ToInt32(IntegerUpDownDispatcherGetValue(Fixed_Time_Value)) : Convert.ToInt32(IntegerUpDownDispatcherGetValue(Stop_Writing_Value));
                    Dictionary<string, string> dict = null;
                    for (int file = 0; file < vsm_file.Length; file++)
                    {
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].Open_Text_val, "點選資料夾圖示"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        SetWindowsPos(VSM_Windows_Title, VSM_Windows_X, VSM_Windows_Y, VSM_Windows_Width, VSM_Windows_Height, "設定vsm對話視窗");
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].ChooseVSMPath_Text_val, "選擇vsm路徑"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        Do.SimulateInputText(TextBoxDispatcherGetValue(VSM_File_Location), "輸入vsm檔路徑");
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].VSM_Text_val, "點選開啟檔案類型欄位"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].VSMType_Text_val, "選擇開檔類型vsm"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].InputVSM_Text_val, "點選輸入vsm檔欄位"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        Do.SimulateInputText(Path.GetFileName(vsm_file[file]), "輸入vsm檔名");
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].TurnOn_Text_val, "點選開啟"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].View_Text_val, "點選view"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].Display_Text_val, "點選Display"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].OnePane_Text_val, "點選1Pane"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].Magnification_Text_val, "點選放大"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        Thread.Sleep(3000);
                        foreach (var point in WaferPointParameter_info[0].WaferPoint_val)
                        {
                            if (!Do.SimulateLeftMouseDoubleClick(Do.ConvertWaferCoordStr(point).Item2, "點選Wafer點位"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].DTCS_Text_val, "點選DTCS"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].OK_Text_val, "點選OK"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            SetWindowsPosVDSW(vsm_file[file], VDSW_Windows_X, VDSW_Windows_Y, VDSW_Windows_Width, VDSW_Windows_Height);
                            if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].Save_Text_val, "點選Save"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            SetWindowsPos(Dat_Windows_Title, Dat_Windows_X, Dat_Windows_Y, Dat_Windows_Width, Dat_Windows_Height, "設定dat對話視窗");
                            if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].Dat_Text_val, "點選存檔類型欄位"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].DatType_Text_val, "選擇儲存類型dat"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].InputDat_Text_val, "點選輸入dat檔欄位"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            Do.SimulateInputText(Path.GetFileNameWithoutExtension(vsm_file[file]) + Do.ConvertWaferCoordStr(point).Item1, "輸入dat檔名");
                            if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].Archive_Text_val, "點選存檔"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].CloseVDSW_Text_val, "關閉VDSW"))
                            {
                                cts.Cancel();
                                Do.ForegroundAutoflow();
                                this.Dispatcher.Invoke(() =>
                                {
                                    Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                                });
                                if (cts.Token.IsCancellationRequested)
                                {
                                    return;
                                }
                            }
                            Thread.Sleep(100);
                        }
                        if (!Do.SimulateLeftMouseClick(Step1Parameter_info[0].CloseWafer_Text_val, "關閉Wafer"))
                        {
                            cts.Cancel();
                            Do.ForegroundAutoflow();
                            this.Dispatcher.Invoke(() =>
                            {
                                Logger.WriteLog("有人員操作滑鼠!", LogLevel.General, richTextBoxGeneral);
                            });
                            if (cts.Token.IsCancellationRequested)
                            {
                                return;
                            }
                        }
                        dict = SetCsvWorkPath(vsm_file[file], date);
                        if (Do.MoveFileToUpper(TextBoxDispatcherGetValue(VSM_File_Location), dict["sample_spectrum"], "*dat"))
                        {
                            Do.RunSoftware(TextBoxDispatcherGetValue(Ref_Fit_Location));
                        }
                        if (Do.CheckCSV(dict["WaveCsvPath"], dict["ParameterCsvPath"], method, timeout))
                        {
                            EH.waferID = Path.GetFileNameWithoutExtension(vsm_file[file]);
                            if (CheckBoxDispatcherGetValue(Save_Datfile))
                            {
                                Do.MoveDatFile(dict["sample_spectrum"], Path.Combine(TextBoxDispatcherGetValue(Xlsx_File_Location), date, Path.GetFileName(vsm_file[file])));
                            }
                            else
                            {
                                Do.DeleteFile(dict["sample_spectrum"], "*dat");
                            }
                            if (EH.WaveToScatterChart(dict["WaveCsvPath"], dict["WaveXlsxPath"]))
                            {
                                File.Move(dict["WaveCsvPath"], dict["MoveWaveCsvPath"]);
                            }
                            if (EH.ParameterToScatterChart(dict["ParameterCsvPath"], dict["ParameterXlsxPath"]))
                            {
                                File.Move(dict["ParameterCsvPath"], dict["MoveParameterCsvPath"]);
                            }
                        }
                    };
                    Do.ForegroundAutoflow();
                    this.Dispatcher.Invoke(() =>
                    {
                        Logger.WriteLog("自動化流程完成!", LogLevel.General, richTextBoxGeneral);
                    });
                }
                else
                {
                    MessageBox.Show("vsm檔資料夾內沒有vsm檔!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }, cts.Token);
        }
        #endregion

        private void CheckSendValueInit()
        {
            Step1Data.CheckSendValueEventHandler1 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data1 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler2 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data2 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler3 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data3 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler4 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data4 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler5 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data5 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler6 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data6 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler7 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data7 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler8 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data8 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler9 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data9 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler10 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data10 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler11 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data11 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler12 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data12 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler13 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data13 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler14 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data14 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler15 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data15 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler16 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data16 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler17 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data17 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler18 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data18 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler19 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data19 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
            Step1Data.CheckSendValueEventHandler20 += (val) =>
            {
                if (val == true)
                {
                    Step1Data.Step1_data20 = Do.ConvertCoordStr(_downPoint, Display_Image);
                }
            };
        }

        private void WriteListBoxContent()
        {
            using (StreamWriter file = new StreamWriter(@"Config/ListBox.txt"))
            {
                foreach (var item in Model_Type_Checklist.Items)
                {
                    file.WriteLine(item.ToString());
                }
            }
        }

        private void ReadListBoxContent()
        {
            string filePath = @"Config/ListBox.txt";
            if (File.Exists(filePath))
            {
                using (StreamReader sr = new StreamReader(filePath))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        Model_Type.Items.Add(line);
                    }
                }
            }
        }
        #endregion

        #region Parameter and Init
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //Lock();
            InitialTray();
            LoadConfig();
            ReadListBoxContent();
            Do.LoadEIM();
            Do.CloseCapsLock();
            CheckSendValueInit();
            Model_Type_state = true;
            Wafer_Type_state = true;
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
        private bool Wafer_Type_state = false;
        List<AutoFlow.StepWindow.Step1Parameter> Step1Parameter_info;
        List<AutoFlow.StepWindow.WaferPointParameter> WaferPointParameter_info;
        CancellationTokenSource cts;
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
                        Logger.WriteLog("擷取影像!", LogLevel.General, richTextBoxGeneral);
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
                            Logger.WriteLog("設定setting檔案路徑!", LogLevel.General, richTextBoxGeneral);
                        }
                        break;
                    }
                case nameof(Open_VSM_Folder):
                    {
                        System.Windows.Forms.FolderBrowserDialog open_vsm_folder_path = new System.Windows.Forms.FolderBrowserDialog();
                        open_vsm_folder_path.Description = "選擇vsm檔資料夾";
                        if (open_vsm_folder_path.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            VSM_File_Location.Text = open_vsm_folder_path.SelectedPath;
                            Logger.WriteLog("設定vsm檔資料夾路徑!", LogLevel.General, richTextBoxGeneral);
                        }
                        break;
                    }
                case nameof(Open_Ref_Fit_Folder):
                    {
                        System.Windows.Forms.FolderBrowserDialog open_vsm_folder_path = new System.Windows.Forms.FolderBrowserDialog();
                        open_vsm_folder_path.Description = "選擇Ref-Fit軟體資料夾";
                        if (open_vsm_folder_path.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            Ref_Fit_Location.Text = open_vsm_folder_path.SelectedPath;
                            Logger.WriteLog("設定Ref-Fit軟體資料夾路徑!", LogLevel.General, richTextBoxGeneral);
                        }
                        break;
                    }
                case nameof(Open_Xlsx_Folder):
                    {
                        System.Windows.Forms.FolderBrowserDialog open_xlsx_folder_path = new System.Windows.Forms.FolderBrowserDialog();
                        open_xlsx_folder_path.Description = "選擇xlsx檔資料夾";
                        if (open_xlsx_folder_path.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            Xlsx_File_Location.Text = open_xlsx_folder_path.SelectedPath;
                            Logger.WriteLog("設定xlsx檔資料夾路徑!", LogLevel.General, richTextBoxGeneral);
                        }
                        break;
                    }
                case nameof(VSM_Set_Windows):
                    {
                        int x = Convert.ToInt32(VSM_Windows_X.Text);
                        int y = Convert.ToInt32(VSM_Windows_Y.Text);
                        int w = Convert.ToInt32(VSM_Windows_Width.Text);
                        int h = Convert.ToInt32(VSM_Windows_Height.Text);
                        Do.SetWindowsPosition(VSM_Windows_Title.Text, new Tuple<int, int, int, int>(x, y, w, h));
                        Logger.WriteLog("設定vsm對話視窗位置!", LogLevel.General, richTextBoxGeneral);
                        break;
                    }
                case nameof(Dat_Set_Windows):
                    {
                        int x = Convert.ToInt32(Dat_Windows_X.Text);
                        int y = Convert.ToInt32(Dat_Windows_Y.Text);
                        int w = Convert.ToInt32(Dat_Windows_Width.Text);
                        int h = Convert.ToInt32(Dat_Windows_Height.Text);
                        Do.SetWindowsPosition(Dat_Windows_Title.Text, new Tuple<int, int, int, int>(x, y, w, h));
                        Logger.WriteLog("設定dat對話視窗位置!", LogLevel.General, richTextBoxGeneral);
                        break;
                    }
                case nameof(VDSW_Set_Windows):
                    {
                        SetWindowsPosVDSW(VDSW_Windows_Title.Text, VDSW_Windows_X, VDSW_Windows_Y, VDSW_Windows_Width, VDSW_Windows_Height);
                        break;
                    }
                case nameof(Save_Config):
                    {
                        SaveConfig();
                        Logger.WriteLog("儲存參數!", LogLevel.General, richTextBoxGeneral);
                        break;
                    }
                case nameof(Add_Item):
                    {
                        string item = Input_Model.Text;
                        if (!string.IsNullOrEmpty(item))
                        {
                            Model_Type_Checklist.Items.Add(item);
                            Logger.WriteLog("增加模型名稱!", LogLevel.General, richTextBoxGeneral);
                        }
                        break;
                    }
                case nameof(Delete_Item):
                    {
                        string item = Input_Model.Text;
                        if (!string.IsNullOrEmpty(item))
                        {
                            Model_Type_Checklist.Items.Remove(item);
                            Logger.WriteLog("刪除模型名稱!", LogLevel.General, richTextBoxGeneral);
                        }
                        break;
                    }
                case nameof(Change_Model_Item):
                    {
                        List<Parameter> Parameter_info = Config.Load();
                        string Content = Parameter_info[0].Model_Type_val;
                        if (!string.IsNullOrEmpty(Content))
                        {
                            Model_Type.Items.Clear();
                            if (!Model_Type_Checklist.Items.Cast<object>().Any(item => item.ToString() == Content))
                            {
                                Model_Type.Items.Add(Content);
                            }
                            Model_Type.SelectedValue = Content;
                            foreach (var item in Model_Type_Checklist.Items)
                            {
                                Model_Type.Items.Add(item.ToString());
                            }
                            WriteListBoxContent();
                            Logger.WriteLog("變更模型下拉按鈕項目!", LogLevel.General, richTextBoxGeneral);
                        }
                        else
                        {
                            Model_Type.Items.Add("D4-1");
                            Model_Type.SelectedValue = "D4-1";
                            SaveConfig();
                        }
                        break;
                    }
            }
        }
        #endregion

        #region Mouse Button Event
        private List<Ellipse> dotlist = new List<Ellipse>();
        private List<Line> cross1list = new List<Line>();
        private List<Line> cross2list = new List<Line>();
        private List<System.Drawing.Point> Screenpointlist = new List<System.Drawing.Point>();
        private List<System.Windows.Point> pointlist = new List<System.Windows.Point>();
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
            if (Wafer_Type_state)
            {
                if (IsWaferWindowVisible(WWInstance))
                {
                    WWInstance.Close();
                }
                if (IsStep1WindowVisible(SWInstance))
                {
                    SWInstance.Close();
                }
                Logger.WriteLog("因切換晶圓尺寸，關閉相關參數頁面，請重新開啟!", LogLevel.General, richTextBoxGeneral);
            }
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
                        if (Model_Type.SelectedValue != null)
                        {
                            Do.CheckModel(Setting_File_Location.Text, Model_Type.SelectedValue.ToString().Replace(stringToRemove, ""));
                            Logger.WriteLog($"更新{Model_Type.SelectedValue.ToString().Replace(stringToRemove, "")}模型!", LogLevel.General, richTextBoxGeneral);
                        }
                    }
                    else
                    {
                        MessageBox.Show("setting檔案不存在!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        Logger.WriteLog("更新模型失敗!原因:setting檔案不存在!", LogLevel.General, richTextBoxGeneral);
                    }
                }
                else
                {
                    MessageBox.Show("請輸入setting檔案位置!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Logger.WriteLog("更新模型失敗!原因:setting檔案路徑欄位為空!", LogLevel.General, richTextBoxGeneral);
                }
            }
        }
        #endregion

    }
}
