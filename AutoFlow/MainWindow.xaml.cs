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

namespace AutoFlow
{
    public class Parameter
    {
        public string Window_Name_val { get; set; }
        public string Coordinate_X_val { get; set; }
        public string Coordinate_Y_val { get; set; }
        public List<System.Drawing.Point> click_points { get; set; }
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
        private void InitialTray()
        {
            notifyIcon = new System.Windows.Forms.NotifyIcon();
            notifyIcon.Icon = new System.Drawing.Icon(@"Icon/Deepwise.ico");
            notifyIcon.Text = "Elf";
            notifyIcon.Visible = true;
            notifyIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(notifyIcon_MouseClick);
            this.StateChanged += new EventHandler(WPFUI_StateChanged);
            //小圖示選單
            nIconMenuItem1.Index = 0;
            nIconMenuItem1.Text = "結束";
            nIconMenuItem1.Click += new System.EventHandler(nIconMenuItem1_Click);
            nIconMenu.MenuItems.Add(nIconMenuItem1);
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
            System.Windows.Application.Current.Shutdown();
        }
        #endregion

        public string TextBoxDispatcherGetValue(TextBox control)
        {
            string name = "";
            this.Dispatcher.Invoke(() =>
            {
                name = control.Text;
            });
            return name;

        }

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

        #region Config
        private void LoadConfig()
        {
            List<Parameter> Parameter_info = Config.Load();
            Window_Name.Text = Parameter_info[0].Window_Name_val;
            Coordinate_X.Text = Parameter_info[0].Coordinate_X_val;
            Coordinate_Y.Text = Parameter_info[0].Coordinate_Y_val; 
        }

        private void SaveConfig()
        {
            List<Parameter> Parameter_config = new List<Parameter>()
            {
                new Parameter() {
                    Window_Name_val = Window_Name.Text,
                    Coordinate_X_val = Coordinate_X.Text,
                    Coordinate_Y_val = Coordinate_Y.Text,
                    click_points = pointlist
                }
            };
            Config.Save(Parameter_config);

        }
        #endregion

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
            canvas.Children.Add(cross1);
            canvas.Children.Add(cross2);
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
            canvas.Children.Add(dot);
        }

        private void GetClickPoint()
        {
            //DrawDot();
            DrawCross();
            // List螢幕座標位置
            int screen_x = Convert.ToInt32(_downPoint.X / Display_Image.ActualWidth * 1920);
            int screen_y = Convert.ToInt32(_downPoint.Y / Display_Image.ActualHeight * 1080);
            pointlist.Add(new System.Drawing.Point(screen_x, screen_y));
            //Console.WriteLine($"螢幕X座標:{screen_x}");
            //Console.WriteLine($"螢幕Y座標:{screen_y}");
        }

        private void RemoveDot()
        {
            if (dotlist.Count != 0)
            {
                canvas.Children.Remove(dotlist[dotlist.Count - 1]);
                dotlist.Remove(dotlist[dotlist.Count - 1]);
                pointlist.Remove(pointlist[pointlist.Count - 1]);
            }
        }

        private void RemoveCross()
        {
            if (cross1list.Count != 0)
            {
                canvas.Children.Remove(cross1list[cross1list.Count - 1]);
                canvas.Children.Remove(cross2list[cross2list.Count - 1]);
                cross1list.Remove(cross1list[cross1list.Count - 1]);
                cross2list.Remove(cross2list[cross2list.Count - 1]);
                pointlist.Remove(pointlist[pointlist.Count - 1]);
            }
        }

        private void RemoveClickPoint()
        {
            //RemoveDot();
            RemoveCross();
        }
        #endregion

        #region Parameter and Init
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            InitialTray();
            LoadConfig();
            //Do.LoadEIM();
            //Do.CloseCapsLock();
        }
        BaseConfig<Parameter> Config = new BaseConfig<Parameter>();
        Core Do = new Core();
        ExcelHandler EH = new ExcelHandler();
        BaseLogRecord Logger = new BaseLogRecord();
        CancellationTokenSource cts;
        private bool _started;
        private System.Windows.Point _downPoint;
        #endregion

        #region Main Screen
        private void Main_Btn_Click(object sender, RoutedEventArgs e)
        {
            switch ((sender as Button).Name)
            {
                case nameof(Start):
                    {
                        cts = new CancellationTokenSource();
                        Task.Run(() =>
                        {
                            while (true)
                            {
                                if (cts.Token.IsCancellationRequested)
                                {
                                    Console.WriteLine("Stop running the software.");
                                    return;
                                }
                                IntPtr targetWindowHandle = Do.PackFindWindow(null, TextBoxDispatcherGetValue(Window_Name));
                                if (targetWindowHandle != IntPtr.Zero)
                                {
                                    #region Get window position and size.
                                    //Console.WriteLine($"找到了 {TextBoxDispatcherGetValue(Window_Name)} 的視窗句柄: {targetWindowHandle}");
                                    //RECT windowRect;
                                    //GetWindowRect(targetWindowHandle, out windowRect);
                                    //Console.WriteLine($"視窗位置: ({windowRect.Left}, {windowRect.Top})");
                                    //Console.WriteLine($"視窗大小: {windowRect.Right - windowRect.Left} x {windowRect.Bottom - windowRect.Top}");
                                    #endregion
                                    Do.PackSetForegroundWindow(targetWindowHandle);
                                    // Action process example:
                                    Do.SimulateRightMouseClick(targetWindowHandle, Convert.ToInt32(TextBoxDispatcherGetValue(Coordinate_X)), Convert.ToInt32(TextBoxDispatcherGetValue(Coordinate_Y)));
                                    System.Windows.Forms.SendKeys.SendWait("D:\\oCam");
                                    //Do.SimulateLeftMouseClick(targetWindowHandle, 899, 156);
                                    Thread.Sleep(3000);
                                }
                                else
                                {
                                    Console.WriteLine($"{TextBoxDispatcherGetValue(Window_Name)} Window can't be found.");
                                }
                                Thread.Sleep(3000);
                            }
                        }, cts.Token);
                        break;
                    }
                case nameof(Stop):
                    {
                        //cts.Cancel();
                        string csvFilePath = @"D:\TEST.csv";
                        EH.CSVToList(csvFilePath, new Tuple<int, int>(1, 2));
                        break;
                    }
                case nameof(Capture_Screen):
                    {
                        Do.CaptureScreen(Display_Image);
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
                case nameof(Save_Config):
                    {
                        SaveConfig();
                        Logger.WriteLog("Save the config.", LogLevel.General, richTextBoxGeneral);
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

    }
}
