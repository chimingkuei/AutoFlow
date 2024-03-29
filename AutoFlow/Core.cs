﻿using AutoFlow;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OpenCvSharp.Extensions;
using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Markup;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using static AutoFlow.MainWindow;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using Newtonsoft.Json;
using OpenCvSharp.Flann;
using System.Windows.Media.TextFormatting;
using System.Collections;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices.ComTypes;
using System.Globalization;
using OpenCvSharp.Internal;

namespace AutoFlow
{
    class Core
    {
        #region Find windows
        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        public IntPtr PackFindWindows(string lpClassName, string lpWindowName)
        {
            return FindWindow(lpClassName, lpWindowName);
        }
        #endregion

        #region Set foreground windows
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindows(IntPtr hWnd);
        public bool PackSetForegroundWindows(IntPtr hWnd)
        {
            return SetForegroundWindows(hWnd);
        }
        #endregion

        #region Set windows position
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        const int SWP_NOSIZE = 0x0001;
        const int SWP_NOMOVE = 0x0002;
        public bool SetWindowsPosition(string windows_title, Tuple<int, int ,int, int> position)
        {
            IntPtr hWnd = FindWindow(null, windows_title);
            if (hWnd != IntPtr.Zero)
            {
                SetWindowPos(hWnd, IntPtr.Zero, position.Item1, position.Item2, position.Item3, position.Item4, 0);
                return true;
            }
            else
            {
                System.Windows.MessageBox.Show($"未找到標題{windows_title}視窗!", "確認", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
        }
        #endregion

        #region Get mouse position
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);
        // 定義POINT結構
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }
        public void GetMousePosition()
        {
            POINT point;
            GetCursorPos(out point);
            Console.WriteLine($"Mouse Position - X: {point.X}, Y: {point.Y}");
        }
        #endregion

        #region Coordinate Format Conversion
        public string ConvertCoordStr(System.Windows.Point point, System.Windows.Controls.Image display_image)
        {
            if (point != new System.Windows.Point(0, 0))
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

        public System.Drawing.Point ConvertCoordXY(string coord_str)
        {
            Match match = Regex.Match(coord_str, @"\((\d+),(\d+)\)");
            return new System.Drawing.Point(int.Parse(match.Groups[1].Value), int.Parse(match.Groups[2].Value));
        }

        public Tuple<string, string> ConvertWaferCoordStr(string coord_str)
        {
            string wafer_coord = "_X" + coord_str.Split(',')[0].Trim('(') + "_Y" + coord_str.Split(',')[1].Trim(')');
            string wafer_screen_coord = "\"" + coord_str.Split(',')[2] + "," + coord_str.Split(',')[3] + "\"";
            return new Tuple<string, string>(wafer_coord, wafer_screen_coord);
        }
        #endregion

        #region Mouse action
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetCursorPos(int x, int y);
        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, IntPtr dwExtraInfo);
        // 定義滑鼠事件的標誌位
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        public bool SimulateLeftMouseClick(string coord_str, string annotation = null)
        {
            System.Drawing.Point pos = ConvertCoordXY(coord_str);
            SetCursorPos(pos.X, pos.Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
            POINT point;
            GetCursorPos(out point);
            return (pos.X - point.X != 0 && pos.Y - point.Y != 0) ? false : true;
        }

        public bool SimulateLeftMouseDoubleClick(string coord_str, string annotation = null)
        {
            System.Drawing.Point pos = ConvertCoordXY(coord_str);
            SetCursorPos(pos.X, pos.Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
            POINT point;
            GetCursorPos(out point);
            return (pos.X - point.X != 0 && pos.Y - point.Y != 0) ? false : true;
        }

        public bool SimulateRightMouseClick(string coord_str, string annotation = null)
        {
            System.Drawing.Point pos = ConvertCoordXY(coord_str);
            SetCursorPos(pos.X, pos.Y);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
            POINT point;
            GetCursorPos(out point);
            return (pos.X - point.X != 0 && pos.Y - point.Y != 0) ? false : true;
        }

        public bool SimulateRightMouseDoubleClick(string coord_str, string annotation = null)
        {
            System.Drawing.Point pos = ConvertCoordXY(coord_str);
            SetCursorPos(pos.X, pos.Y);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
            POINT point;
            GetCursorPos(out point);
            return (pos.X - point.X != 0 && pos.Y - point.Y != 0) ? false : true;
        }

        public void SimulateInputText(string keys, string annotation = null)
        {
            System.Windows.Forms.SendKeys.SendWait(keys);
        }
        #endregion

        #region Close caps lock
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        const int KEYEVENTF_EXTENDEDKEY = 0x1;

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
        public static extern short GetKeyState(int keyCode);

        public void CloseCapsLock()
        {
            bool CapsLock = (((ushort)GetKeyState(0x14)) & 0xffff) != 0;
            if (CapsLock)
                keybd_event(0x14, 0x45, KEYEVENTF_EXTENDEDKEY, (UIntPtr)0);
        }
        #endregion

        #region Load english input method
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint Flags);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr ActivateKeyboardLayout(IntPtr hkl, uint Flags);
        // 定義輸入法標識符號
        const string ENGLISH_KEYBOARD_LAYOUT_ID = "00000409"; // 英文（美國）
        // 定義激活輸入法標誌
        const uint KLF_ACTIVATE = 1;
        public void LoadEIM()
        {
            IntPtr englishLayout = LoadKeyboardLayout(ENGLISH_KEYBOARD_LAYOUT_ID, 0);
            if (englishLayout == IntPtr.Zero)
            {
                Console.WriteLine("載入輸入法失敗!");
                return;
            }
            IntPtr result = ActivateKeyboardLayout(englishLayout, KLF_ACTIVATE);
            if (result == IntPtr.Zero)
            {
                Console.WriteLine("激活輸入法失敗!");
                return;
            }
        }
        #endregion

        #region IO operation
        public string[] GetFilename(string folderPath, string filetype)
        {
            return Directory.GetFiles(folderPath, filetype);
        }

        public bool MoveFileToUpper(string sourceDirectory, string targetDirectory, string type)
        {
            string[] datFiles = Directory.GetFiles(sourceDirectory, type);
            foreach (string datFile in datFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(datFile);
                string newFileName = fileName.ToUpper() + "." + type.Trim('*');
                string targetPath = Path.Combine(targetDirectory, newFileName);
                File.Move(datFile, targetPath);
            }
            return true;
        }

        public void DeleteFile(string sourceDirectory, string type)
        {
            string[] datFiles = Directory.GetFiles(sourceDirectory, type);
            foreach (string datFile in datFiles)
            {
                File.Delete(datFile);
            }
        }

        public void MoveDatFile(string sourceDirectory, string targetDirectory)
        {
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
                Thread.Sleep(200);
            }
            string[] datFiles = Directory.GetFiles(sourceDirectory, "*dat");
            foreach (string datFile in datFiles)
            {
                File.Move(datFile, Path.Combine(targetDirectory, Path.GetFileName(datFile)));
            }
        }

        public bool CheckCSV(string csvfile1, string csvfile2, string type, int timeout)
        {
            bool state = false;
            switch (type)
            {
                case "FixedTime":
                    {
                        while (true)
                        {
                            if (File.Exists(csvfile1) && File.Exists(csvfile2))
                            {
                                state = true;
                                break;
                            }
                            Thread.Sleep(1000);
                        }
                        Thread.Sleep(timeout*1000);
                        break;
                    }
                case "StopWriting":
                    {
                        while (true)
                        {
                            if (File.Exists(csvfile1) && File.Exists(csvfile2))
                            {
                                DateTime lastWriteTime = File.GetLastWriteTime(csvfile2);
                                if (DateTime.Now - lastWriteTime > TimeSpan.FromSeconds(timeout))
                                {
                                    state = true;
                                    break;
                                }
                            }
                            Thread.Sleep(1000);
                        }
                        break;
                    }
            }
            return state;
        }
        #endregion

        private BitmapImage ConvertBitmapToBitmapImage(Bitmap bitmap)
        {
            using (var memory = new System.IO.MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }

        public void CaptureScreen(System.Windows.Controls.Image display_image)
        {
            Rectangle screenBounds = Screen.PrimaryScreen.Bounds;
            Bitmap screenshot = new Bitmap(screenBounds.Width, screenBounds.Height, PixelFormat.Format32bppArgb);

            using (Graphics graphics = Graphics.FromImage(screenshot))
            {
                graphics.CopyFromScreen(screenBounds.X, screenBounds.Y, 0, 0, screenBounds.Size, CopyPixelOperation.SourceCopy);
            }
            BitmapImage bitmapImage = ConvertBitmapToBitmapImage(screenshot);
            display_image.Source = bitmapImage;
        }

        public void CheckModel(string filePath, string model)
        {
            string jsonContent = File.ReadAllText(filePath);
            JObject jsonObject = JObject.Parse(jsonContent);
            jsonObject["design"] = model;
            File.WriteAllText(filePath, jsonObject.ToString());
        }

        public void RunSoftware(string softwarepath)
        {
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.Start();
            process.StandardInput.WriteLine("cd "+ softwarepath);
            process.StandardInput.WriteLine("run_RefFitTool.exe");
            process.Close();
        }
    }

    class ExcelHandler
    {
        public string waferID { get; set; }

        #region Generate output_waveform.xlsx and output_parameters.xlsx commonly
        private bool CheckChartName(ExcelWorksheet worksheet, string chartname)
        {
            foreach (ExcelDrawing drawing in worksheet.Drawings)
            {
                if (drawing.Name == chartname)
                {
                    return true;
                }
            }
            return false;
        }

        private ExcelRange GetRange(ExcelWorksheet worksheet, string field, int start, int end)
        {
            return worksheet.Cells[field + (start).ToString() + ":" + field + (end).ToString()];
        }
        #endregion

        #region  Generate output_waveform.xlsx
        private Dictionary<string, int> WaveGetTupleExtremum(List<List<Tuple<string, string, double, double>>> lists)
        {
            double XmaxTuple0 = lists.SelectMany(innerList => innerList).Max(tuple => tuple.Item3);
            double XminTuple0 = lists.SelectMany(innerList => innerList).Min(tuple => tuple.Item3);
            double YmaxTuple1 = lists.SelectMany(innerList => innerList).Max(tuple => tuple.Item4);
            double YminTuple1 = lists.SelectMany(innerList => innerList).Min(tuple => tuple.Item4);
            Dictionary<string, int> dict = new Dictionary<string, int>();
            dict["XMax"] = Convert.ToInt32(XmaxTuple0) + 50;
            dict["XMin"] = Convert.ToInt32(XminTuple0) - 50;
            dict["YMax"] = Convert.ToInt32(YmaxTuple1);
            dict["YMin"] = Convert.ToInt32(YminTuple1);
            return dict;
        }

        private void WaveSetChartStyle(ExcelChart chart, Tuple<int, int, int, int> position, List<List<Tuple<string, string, double, double>>> lists, string title)
        {
            Dictionary<string, int> range= WaveGetTupleExtremum(lists);
            chart.SetPosition(position.Item1, position.Item2, position.Item3, position.Item4);
            chart.SetSize(600, 400);
            chart.Title.Text = title;
            chart.Legend.Position = eLegendPosition.Right;
            chart.XAxis.MajorGridlines.Fill.Color = Color.LightGray;
            chart.XAxis.MaxValue = range["XMax"];
            chart.XAxis.MinValue = range["XMin"];
            chart.YAxis.MajorGridlines.Fill.Color = Color.LightGray;
            chart.YAxis.MaxValue = range["YMax"];
            chart.YAxis.MinValue = range["YMin"];
        }

        private void WaveFieldLabel(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1"].Value = "tag";
            worksheet.Cells["B1"].Value = "filename";
            worksheet.Cells["C1"].Value = "wl";
            worksheet.Cells["D1"].Value = "amp";
        }

        private void WaveWhiteCells(ExcelWorksheet worksheet, List<List<Tuple<string, string, double, double>>> lists, int tag_count, int cell_init, int list_group)
        {
            for (int cell_index = 0; cell_index < tag_count; cell_index++)
            {
                worksheet.Cells["A" + (cell_init + cell_index).ToString()].Value = lists[list_group][cell_index].Item1;
                worksheet.Cells["B" + (cell_init + cell_index).ToString()].Value = lists[list_group][cell_index].Item2;
                worksheet.Cells["C" + (cell_init + cell_index).ToString()].Value = lists[list_group][cell_index].Item3;
                worksheet.Cells["D" + (cell_init + cell_index).ToString()].Value = lists[list_group][cell_index].Item4;
            }
        }

        private List<List<Tuple<string, string, double, double>>> CSVToList(string csvfilepath)
        {
            List<List<Tuple<string, string, double, double>>> dataListChunks = new List<List<Tuple<string, string, double, double>>>();
            if (File.Exists(csvfilepath))
            {
                using (StreamReader reader = new StreamReader(csvfilepath))
                {
                    // 跳過標題行
                    reader.ReadLine();
                    List<Tuple<string, string, double, double>> currentChunk = new List<Tuple<string, string, double, double>>();
                    string tmp = "0";
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] fields = line.Split(',');
                        if (fields[0] == tmp)
                        {
                            Tuple<string, string, double, double> rowData = new Tuple<string, string, double, double>(fields[0], fields[1], Convert.ToDouble(fields[2]), Convert.ToDouble(fields[3]));
                            currentChunk.Add(rowData);
                        }
                        else
                        {
                            dataListChunks.Add(currentChunk);
                            tmp = fields[0];
                            currentChunk = new List<Tuple<string, string, double, double>>();
                            Tuple<string, string, double, double> rowData = new Tuple<string, string, double, double>(fields[0], fields[1], Convert.ToDouble(fields[2]), Convert.ToDouble(fields[3]));
                            currentChunk.Add(rowData);
                        }
                    }
                    dataListChunks.Add(currentChunk);
                }
                #region For debug
                //foreach (var chunk in dataListChunks)
                //{
                //    Console.WriteLine("--------------------------");
                //    foreach (var tuple in chunk)
                //    {
                //        Console.WriteLine($"({tuple.Item1}, {tuple.Item2}, {tuple.Item3}, {tuple.Item4})");
                //    }
                //}
                #endregion
            }
            return dataListChunks;
        }

        public bool WaveToScatterChart(string csvfilepath, string xlsxfilepath)
        {
            List<List<Tuple<string, string, double, double>>> lists = CSVToList(csvfilepath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("output_waveform");
                WaveFieldLabel(worksheet);
                int cell_y = 2;
                for (int list_index = 0; list_index < lists.Count; list_index += 2)
                {
                    string chartname = "ScatterPlot" + list_index.ToString();
                    if (!CheckChartName(worksheet, chartname))
                    {
                        ExcelChart chart = worksheet.Drawings.AddChart(chartname, eChartType.XYScatterLinesNoMarkers);
                        WaveSetChartStyle(chart, new Tuple<int, int, int, int>((list_index/2)*20, 0, 5, 0), lists, Path.GetFileNameWithoutExtension(lists[list_index][0].Item2));
                        //WaveSetChartStyle(chart, new Tuple<int, int, int, int>(cell_y, 0, 5, 0), lists);
                        int tag0_count = lists[list_index].Count;
                        int tag1_count = lists[list_index + 1].Count;
                        WaveWhiteCells(worksheet, lists, tag0_count, cell_y, list_index);
                        WaveWhiteCells(worksheet, lists, tag1_count, cell_y + tag0_count, list_index + 1);
                        var measurementA = GetRange(worksheet, "C", cell_y, cell_y + tag0_count - 1);
                        var measurementB = GetRange(worksheet, "D", cell_y, cell_y + tag0_count - 1);
                        var simulationA = GetRange(worksheet, "C", cell_y + tag0_count, cell_y + tag0_count + tag1_count - 1);
                        var simulationB = GetRange(worksheet, "D", cell_y + tag0_count, cell_y + tag0_count + tag1_count - 1);
                        var measurementseries = (ExcelScatterChartSerie)chart.Series.Add(measurementB, measurementA);
                        var simulationseries = (ExcelScatterChartSerie)chart.Series.Add(simulationB, simulationA);
                        measurementseries.Header = "0-量測";
                        simulationseries.Header = "模擬";
                        cell_y += tag0_count + tag1_count;
                        //simulationseries.Marker.Style = eMarkerStyle.Star;
                    }
                }
                try
                {
                    package.SaveAs(new FileInfo(xlsxfilepath));
                }
                catch
                {
                    Console.WriteLine(xlsxfilepath + " file is opened! Please close that file.");
                }

            }
            return true;
        }
        #endregion

        #region Generate output_parameters.xlsx
        private void ParameterFieldLabel(ExcelWorksheet worksheet)
        {
            string[] labels = { "tag", "filename", "X", "Y", "", "Gp1", "Gp2", "Gp3", "Gp4", "WL", "gth", "dn", "dip(meas)", "sbc(meas)", "sbw(meas)", "dip(sim)", "sbc(sim)", "sbw(sim)", "Gp1", "Gp2", "Gp3"};
            for (int i = 0; i < labels.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = labels[i];
            }
        }

        private List<string> ParameterCSVToList(string csvfilepath)
        {
            List<string> currentChunk = new List<string>();
            if (File.Exists(csvfilepath))
            {
                using (StreamReader reader = new StreamReader(csvfilepath))
                {
                    // 跳過標題行
                    reader.ReadLine();
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        currentChunk.Add(line);
                    }
                }
            }
            return currentChunk;
        }

        private void ParameterSetChartStyle(ExcelChart chart, Tuple<int, int, int, int> position, string title)
        {
            chart.SetPosition(position.Item1, position.Item2, position.Item3, position.Item4);
            chart.SetSize(600, 400);
            if (title == "Gp1")
            {
                chart.Title.Text = waferID + "_P-DBR";
            }
            else if (title == "Gp2")
            {
                chart.Title.Text = waferID + "_Cavity thickness";
            }
            else if (title == "Gp3")
            {
                chart.Title.Text = waferID + "_N-DBR";
            }
            else
            {
                chart.Title.Text = "Gp1";
            }
            chart.Legend.Position = eLegendPosition.Right;
            chart.XAxis.MajorGridlines.Fill.Color = Color.LightGray;
            chart.XAxis.MajorUnit = 10;
            chart.XAxis.Title.Text = "Away From center (mm)";
            chart.YAxis.MajorGridlines.Fill.Color = Color.LightGray;
            chart.YAxis.Title.TextVertical = eTextVerticalType.Vertical270;
            if (title == "Gp2")
            {
                chart.YAxis.Title.Text = "Shifted Thickness (mm)";
            }
            else
            {
                chart.YAxis.Title.Text = "Shifted %";
            }
            
        }

        private int GetCoord(string filename, string type)
        {
            int coord = 0;
            if (type == "X")
            {
                coord = Convert.ToInt32(Path.GetFileNameWithoutExtension(filename).Split('_')[1].Trim('X'));
            }
            else if (type == "Y")
            {
                coord = Convert.ToInt32(Path.GetFileNameWithoutExtension(filename).Split('_')[2].Trim('Y'));
            }
            return coord;
        }

        private int GetNonZeroCoord(string filename)
        {
            int coord = Convert.ToInt32(Path.GetFileNameWithoutExtension(filename).Split('_')[1].Trim('X'));
            if (coord == 0)
            {
                coord = Convert.ToInt32(Path.GetFileNameWithoutExtension(filename).Split('_')[2].Trim('Y'));
            }
            return coord;
        }

        private Dictionary<string, int> FindZeroIndex(List<string> list_string)
        {
            List<int> coord_x = new List<int>();
            for (int i = 0; i < list_string.Count; i++)
            {
                string filename = list_string[i].Split(',')[1];
                int index = GetCoord(filename, "X");
                coord_x.Add(index);
            }
            int firstZeroIndex = coord_x.FindIndex(num => num == 0);
            int lastZeroIndex = coord_x.FindLastIndex(num => num == 0);
            Dictionary<string, int> dict = new Dictionary<string, int>();
            dict.Add("firstZeroIndex", firstZeroIndex);
            dict.Add("lastZeroIndex", lastZeroIndex);
            return dict;
        }

        private void DrawParameterScatterChart(ExcelWorksheet worksheet, List<string> list_string, string chartnameGp, int pos, string title, string field, int yaxis_length)
        {
            if (!CheckChartName(worksheet, chartnameGp))
            {
                ExcelChart chart = worksheet.Drawings.AddChart(chartnameGp, eChartType.XYScatter);
                ParameterSetChartStyle(chart, new Tuple<int, int, int, int>(pos, 0, 22, 0), title);
                var yaxis_x = GetRange(worksheet, "D", 2, yaxis_length + 1);
                var yaxis_y = GetRange(worksheet, field, 2, yaxis_length + 1);
                var yaxis_series = (ExcelScatterChartSerie)chart.Series.Add(yaxis_y, yaxis_x);
                yaxis_series.Marker.Style = eMarkerStyle.Diamond;
                yaxis_series.Fill.Color = Color.AliceBlue;
                yaxis_series.Header = "Y";
                var xaxis_x1 = GetRange(worksheet, "C", yaxis_length + 2, list_string.Count + 1);
                var xaxis_y1 = GetRange(worksheet, field, yaxis_length + 2, list_string.Count + 1);
                var xaxis_series1 = (ExcelScatterChartSerie)chart.Series.Add(xaxis_y1, xaxis_x1);
                xaxis_series1.Marker.Style = eMarkerStyle.Square;
                xaxis_series1.Fill.Color = Color.Orange;
                xaxis_series1.Header = "X";
            }
        }

        private void DrawParameterScatterChart1(ExcelWorksheet worksheet, List<string> list_string, string chartnameGp, int pos, string title)
        {
            if (!CheckChartName(worksheet, chartnameGp))
            {
                ExcelChart chart = worksheet.Drawings.AddChart(chartnameGp, eChartType.XYScatter);
                ParameterSetChartStyle(chart, new Tuple<int, int, int, int>(pos, 0, 22, 0), title);
                var yaxis_x = GetRange(worksheet, "E", 2, list_string.Count + 1);
                var yaxis_y = GetRange(worksheet, "S", 2, list_string.Count + 1);
                var yaxis_series = (ExcelScatterChartSerie)chart.Series.Add(yaxis_y, yaxis_x);
                yaxis_series.Marker.Style = eMarkerStyle.Circle;
                yaxis_series.Fill.Color = Color.AliceBlue;
            }
        }

        public bool ParameterToScatterChart(string csvfilepath, string xlsxfilepath)
        {
            List<string> list_string = ParameterCSVToList(csvfilepath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("output_parameters");
                ParameterFieldLabel(worksheet);
                Dictionary<string, int> dict = FindZeroIndex(list_string);
                int index1 = 2;
                for (int list_string_index = dict["firstZeroIndex"]; list_string_index < list_string.Count; list_string_index++)
                {
                    string[] fields = list_string[list_string_index].Split(',');
                    worksheet.Cells["A" + index1.ToString()].Value = Convert.ToDouble(fields[0]);
                    worksheet.Cells["B" + index1.ToString()].Value = fields[1];
                    worksheet.Cells["C" + index1.ToString()].Value = GetCoord(fields[1], "X");
                    worksheet.Cells["D" + index1.ToString()].Value = GetCoord(fields[1], "Y");
                    worksheet.Cells["E" + index1.ToString()].Value = GetNonZeroCoord(fields[1]);
                    worksheet.Cells["F" + index1.ToString()].Value = Convert.ToDouble(fields[2]);
                    worksheet.Cells["G" + index1.ToString()].Value = Convert.ToDouble(fields[3]);
                    worksheet.Cells["H" + index1.ToString()].Value = Convert.ToDouble(fields[4]);
                    worksheet.Cells["I" + index1.ToString()].Value = Convert.ToDouble(fields[5]);
                    worksheet.Cells["J" + index1.ToString()].Value = Convert.ToDouble(fields[6]);
                    worksheet.Cells["K" + index1.ToString()].Value = Convert.ToDouble(fields[7]);
                    worksheet.Cells["L" + index1.ToString()].Value = Convert.ToDouble(fields[8]);
                    worksheet.Cells["M" + index1.ToString()].Value = Convert.ToDouble(fields[9]);
                    worksheet.Cells["N" + index1.ToString()].Value = Convert.ToDouble(fields[10]);
                    worksheet.Cells["O" + index1.ToString()].Value = Convert.ToDouble(fields[11]);
                    worksheet.Cells["P" + index1.ToString()].Value = Convert.ToDouble(fields[12]);
                    worksheet.Cells["Q" + index1.ToString()].Value = Convert.ToDouble(fields[13]);
                    worksheet.Cells["R" + index1.ToString()].Value = Convert.ToDouble(fields[14]);
                    worksheet.Cells["S" + index1.ToString()].Value = (Convert.ToDouble(fields[2]) - 1) * 100;
                    worksheet.Cells["T" + index1.ToString()].Value = Convert.ToDouble(fields[3]);
                    worksheet.Cells["U" + index1.ToString()].Value = (Convert.ToDouble(fields[4]) - 1) * 100;
                    index1 += 1;
                }
                int index2 = list_string.Count - dict["firstZeroIndex"] + 2;
                for (int list_string_index = 0; list_string_index < dict["firstZeroIndex"]; list_string_index++)
                {
                    string[] fields = list_string[list_string_index].Split(',');
                    worksheet.Cells["A" + index2.ToString()].Value = Convert.ToDouble(fields[0]);
                    worksheet.Cells["B" + index2.ToString()].Value = fields[1];
                    worksheet.Cells["C" + index2.ToString()].Value = GetCoord(fields[1], "X");
                    worksheet.Cells["D" + index2.ToString()].Value = GetCoord(fields[1], "Y");
                    worksheet.Cells["E" + index2.ToString()].Value = GetNonZeroCoord(fields[1]);
                    worksheet.Cells["F" + index2.ToString()].Value = Convert.ToDouble(fields[2]);
                    worksheet.Cells["G" + index2.ToString()].Value = Convert.ToDouble(fields[3]);
                    worksheet.Cells["H" + index2.ToString()].Value = Convert.ToDouble(fields[4]);
                    worksheet.Cells["I" + index2.ToString()].Value = Convert.ToDouble(fields[5]);
                    worksheet.Cells["J" + index2.ToString()].Value = Convert.ToDouble(fields[6]);
                    worksheet.Cells["K" + index2.ToString()].Value = Convert.ToDouble(fields[7]);
                    worksheet.Cells["L" + index2.ToString()].Value = Convert.ToDouble(fields[8]);
                    worksheet.Cells["M" + index2.ToString()].Value = Convert.ToDouble(fields[9]);
                    worksheet.Cells["N" + index2.ToString()].Value = Convert.ToDouble(fields[10]);
                    worksheet.Cells["O" + index2.ToString()].Value = Convert.ToDouble(fields[11]);
                    worksheet.Cells["P" + index2.ToString()].Value = Convert.ToDouble(fields[12]);
                    worksheet.Cells["Q" + index2.ToString()].Value = Convert.ToDouble(fields[13]);
                    worksheet.Cells["R" + index2.ToString()].Value = Convert.ToDouble(fields[14]);
                    worksheet.Cells["S" + index2.ToString()].Value = (Convert.ToDouble(fields[2]) - 1) * 100;
                    worksheet.Cells["T" + index2.ToString()].Value = Convert.ToDouble(fields[3]);
                    worksheet.Cells["U" + index2.ToString()].Value = (Convert.ToDouble(fields[4]) - 1) * 100;
                    index2 += 1;
                }
                DrawParameterScatterChart1(worksheet, list_string, "None", 0, "None");
                DrawParameterScatterChart(worksheet, list_string, "Gp1", 22, "Gp1", "S", dict["lastZeroIndex"] - dict["firstZeroIndex"] + 1);
                DrawParameterScatterChart(worksheet, list_string, "Gp2", 44, "Gp2", "T", dict["lastZeroIndex"] - dict["firstZeroIndex"] + 1);
                DrawParameterScatterChart(worksheet, list_string, "Gp3", 66, "Gp3", "U", dict["lastZeroIndex"] - dict["firstZeroIndex"] + 1);
                try
                {
                    package.SaveAs(new FileInfo(xlsxfilepath));
                }
                catch
                {
                    Console.WriteLine(xlsxfilepath + " file is opened! Please close that file.");
                }
                return true;
            }
        }
        #endregion

        public string ConvertWaferPointJsonFormat(string[] fields)
        {
            return "(" + fields[0] + "," + fields[1] + ")" + "," + "(" + fields[2] + "," + fields[3] + ")";
        }

        public List<string> ReadCsv(string csvfilepath, Func<string[], string> fun)
        {
            List<string> data = new List<string>();
            if (!string.IsNullOrEmpty(csvfilepath))
            {
                if (File.Exists(csvfilepath))
                {
                    using (StreamReader reader = new StreamReader(csvfilepath))
                    {
                        reader.ReadLine();
                        while (!reader.EndOfStream)
                        {
                            string line = reader.ReadLine();
                            string[] fields = line.Split(',');
                            data.Add(fun(fields));
                        }
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("晶圓點位csv檔不存在!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("請輸入晶圓點位csv檔位置!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            return data;
        }

        public List<string> ConvertScreenCoordinate(string csvfilepath, Tuple<int, int> origin_screen_index, double mag)
        {
            List<string> data = new List<string>();
            data.Add("X,Y,Screen X,Screen Y");
            if (!string.IsNullOrEmpty(csvfilepath))
            {
                if (File.Exists(csvfilepath))
                {
                    using (StreamReader reader = new StreamReader(csvfilepath))
                    {
                        reader.ReadLine();
                        while (!reader.EndOfStream)
                        {
                            string line = reader.ReadLine();
                            string[] fields = line.Split(',');
                            string x = Math.Round(origin_screen_index.Item1 + Convert.ToInt32(fields[0]) * mag).ToString();
                            string y = Math.Round(origin_screen_index.Item2 - Convert.ToInt32(fields[1]) * mag).ToString();
                            data.Add(fields[0] + "," + fields[1] + "," + x + "," + y);
                        }
                    }
                    string[] columnData = data.ToArray();
                    using (StreamWriter writer = new StreamWriter(csvfilepath))
                    {
                        for (int i = 0; i < columnData.Length; i++)
                        {
                            string line = columnData[i];
                            if (i < columnData.Length - 1)
                            {
                                writer.WriteLine(line);
                            }
                            else
                            {
                                writer.Write(line);
                            }
                        }
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("晶圓點位csv檔不存在!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("請輸入晶圓點位csv檔位置!", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            return data;
        }
    }
}
