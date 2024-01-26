using AutoFlow;
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

namespace AutoFlow
{
    class Core
    {
        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, IntPtr dwExtraInfo);

        // 定義滑鼠事件的標誌位
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const uint MOUSEEVENTF_RIGHTUP = 0x0010;

        public IntPtr PackFindWindow(string lpClassName, string lpWindowName)
        {
            return FindWindow(lpClassName, lpWindowName);
        }

        public bool PackSetForegroundWindow(IntPtr hWnd)
        {
            return SetForegroundWindow(hWnd);
        }

        public void SimulateInputText(string keys, string annotation = null)
        {
            System.Windows.Forms.SendKeys.SendWait(keys);
        }

        #region Mouse Action
        public void SimulateLeftMouseClick(System.Drawing.Point pos, string annotation = null)
        {
            SetCursorPos(pos.X, pos.Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
        }

        public void SimulateLeftMouseDoubleClick(System.Drawing.Point pos, string annotation = null)
        {
            SetCursorPos(pos.X, pos.Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
        }

        public void SimulateRightMouseClick(System.Drawing.Point pos, string annotation = null)
        {
            SetCursorPos(pos.X, pos.Y);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
        }

        public void SimulateRightMouseDoubleClick(System.Drawing.Point pos, string annotation = null)
        {
            SetCursorPos(pos.X, pos.Y);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
        }
        #endregion

        #region Close Caps Lock
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

        #region Load English Input Method
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
            string jsonData = File.ReadAllText(filePath);
            JArray jsonArray = JArray.Parse(jsonData);
            foreach (JObject item in jsonArray.Children<JObject>())
            {
                JProperty designProperty = item.Property("design");
                if (designProperty != null)
                {
                    designProperty.Value = model;
                }
            }
            File.WriteAllText(filePath, jsonArray.ToString());
        }

        public string[] GetFilename(string folderPath, string filetype)
        {
            return Directory.GetFiles(folderPath, filetype);
        }

        public void RunSoftware(string softwarepath)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = softwarepath,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            Process.Start(startInfo);
        }

        public void MoveDatFile(string sourceDirectory, string targetDirectory)
        {
            string[] datFiles = Directory.GetFiles(sourceDirectory, "*.dat");
            foreach (string datFile in datFiles)
            {
                string fileName = System.IO.Path.GetFileName(datFile);
                string targetPath = System.IO.Path.Combine(targetDirectory, fileName);
                File.Move(datFile, targetPath);
            }
        }

        public void CheckCSV(string csvfile1, string csvfile2)
        {
            while (true)
            {
                if (File.Exists(csvfile1) & File.Exists(csvfile2))
                {
                    break; 
                }
                Thread.Sleep(1000);
            }

        }
    }

    class ExcelHandler
    {
        public string waferID { get; set; }

        private Dictionary<string, int> GetTupleExtremum(List<List<Tuple<double, double>>> lists)
        {
            double XmaxTuple0 = lists.SelectMany(list => list).Max(tuple => tuple.Item1);
            double XminTuple0 = lists.SelectMany(list => list).Min(tuple => tuple.Item1);
            double YmaxTuple1 = lists.SelectMany(list => list).Max(tuple => tuple.Item2);
            double YminTuple1 = lists.SelectMany(list => list).Min(tuple => tuple.Item2);
            Dictionary<string, int> dict = new Dictionary<string, int>();
            dict["XMax"] = Convert.ToInt32(XmaxTuple0) + 50;
            dict["XMin"] = Convert.ToInt32(XminTuple0) - 50;
            dict["YMax"] = Convert.ToInt32(YmaxTuple1);
            dict["YMin"] = Convert.ToInt32(YminTuple1);
            return dict;
        }

        private void SetChartStyle(ExcelChart chart, Tuple<int, int, int, int> position, List<List<Tuple<double, double>>> lists)
        {
            Dictionary<string, int> range= GetTupleExtremum(lists);
            chart.SetPosition(position.Item1, position.Item2, position.Item3, position.Item4);
            chart.SetSize(600, 400);
            chart.Title.Text = waferID;
            chart.Legend.Position = eLegendPosition.Right;
            chart.XAxis.MajorGridlines.Fill.Color = Color.LightGray;
            chart.XAxis.MaxValue = range["XMax"];
            chart.XAxis.MinValue = range["XMin"];
            chart.YAxis.MajorGridlines.Fill.Color = Color.LightGray;
            chart.YAxis.MaxValue = range["YMax"];
            chart.YAxis.MinValue = range["YMin"];
        }

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

        public List<List<Tuple<double, double>>> CSVToList(string csvfilepath, Tuple<int, int, int> index)
        {
            List<List<Tuple<double, double>>> dataListChunks = new List<List<Tuple<double, double>>>();
            if (File.Exists(csvfilepath))
            {
                using (StreamReader reader = new StreamReader(csvfilepath))
                {
                    // 跳過標題行
                    reader.ReadLine();
                    List<Tuple<double, double>> currentChunk = new List<Tuple<double, double>>();
                    string tmp = "0";
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] fields = line.Split(',');
                        if (fields[index.Item1] == tmp)
                        {
                            Tuple<double, double> rowData = new Tuple<double, double>(Convert.ToDouble(fields[index.Item2]), Convert.ToDouble(fields[index.Item3]));
                            currentChunk.Add(rowData);
                        }
                        else
                        {
                            Tuple<double, double> rowData = new Tuple<double, double>(Convert.ToDouble(fields[index.Item2]), Convert.ToDouble(fields[index.Item3]));
                            currentChunk.Add(rowData);
                            dataListChunks.Add(currentChunk);
                            tmp = fields[index.Item1];
                            currentChunk = new List<Tuple<double, double>>();
                        }
                    }
                    dataListChunks.Add(currentChunk);
                }
                #region For debug
                //foreach (var chunk in dataListChunks)
                //{
                //    foreach (var tuple in chunk)
                //    {
                //        Console.WriteLine($"({tuple.Item1}, {tuple.Item2})");
                //    }
                //}
                #endregion
            }
            return dataListChunks;
        }

        // The name of chart can't be repeated.(ScatterPlot)
        public void WaveToScatterChart(string filepath, string sheetname, List<List<Tuple<double, double>>> lists)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetname);
                int cell_y = 0;
                for (int list_index = 0; list_index < lists.Count; list_index += 2)
                {
                    string chartname = "ScatterPlot" + list_index.ToString();
                    if (!CheckChartName(worksheet, chartname))
                    {
                        ExcelChart chart = worksheet.Drawings.AddChart(chartname, eChartType.XYScatterLinesNoMarkers);
                        SetChartStyle(chart, new Tuple<int, int, int, int>(cell_y, 0, list_index + 5, 0), lists);
                        // Set X and Y axis data ranges
                        int tag0_count = lists[list_index].Count;
                        int tag1_count = lists[list_index + 1].Count;
                        for (int cell_index = 0; cell_index < tag0_count; cell_index++)
                        {
                            worksheet.Cells["A" + (cell_y + cell_index + 1).ToString()].Value = lists[list_index][cell_index].Item1;
                            worksheet.Cells["B" + (cell_y + cell_index + 1).ToString()].Value = lists[list_index][cell_index].Item2;
                        }
                        for (int cell_index = 0; cell_index < tag1_count; cell_index++)
                        {
                            worksheet.Cells["A" + (cell_y + tag0_count + cell_index + 1).ToString()].Value = lists[list_index + 1][cell_index].Item1;
                            worksheet.Cells["B" + (cell_y + tag0_count + cell_index + 1).ToString()].Value = lists[list_index + 1][cell_index].Item2;
                        }
                        var measurementA = worksheet.Cells["A" + (cell_y + 1).ToString() + ":" + "A" + (cell_y + tag0_count + 1).ToString()];
                        var measurementB = worksheet.Cells["B" + (cell_y + 1).ToString() + ":" + "B" + (cell_y + tag0_count + 1).ToString()];
                        var simulationA = worksheet.Cells["A" + (cell_y + tag0_count + 1).ToString() + ":" + "A" + (cell_y + tag0_count + tag1_count + 1).ToString()];
                        var simulationB = worksheet.Cells["B" + (cell_y + tag0_count + 1).ToString() + ":" + "B" + (cell_y + tag0_count + tag1_count + 1).ToString()];
                        var measurementseries = (ExcelScatterChartSerie)chart.Series.Add(measurementB, measurementA);
                        var simulationseries = (ExcelScatterChartSerie)chart.Series.Add(simulationB, simulationA);
                        measurementseries.Header = "0-量測";
                        simulationseries.Header = "模擬";
                        cell_y += tag0_count + tag1_count;
                    }
                }
                try
                {
                    package.SaveAs(new FileInfo(filepath));
                }
                catch
                {
                    Console.WriteLine(filepath + " file is opened! Please close that file.");
                }

            }
        }

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

        public List<string> ConvertScreenCoordinate(string csvfilepath, Tuple<int, int> origin_screen_index)
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
                            string x = Math.Round(origin_screen_index.Item1 + Convert.ToInt32(fields[0]) * 5.2).ToString();
                            string y = Math.Round(origin_screen_index.Item2 - Convert.ToInt32(fields[1]) * 5.2).ToString();
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
