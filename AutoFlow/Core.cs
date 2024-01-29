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
using OpenCvSharp.Flann;

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
            string jsonContent = File.ReadAllText(filePath);
            JObject jsonObject = JObject.Parse(jsonContent);
            jsonObject["design"] = model;
            File.WriteAllText(filePath, jsonObject.ToString());
        }

        public string[] GetFilename(string folderPath, string filetype)
        {
            return Directory.GetFiles(folderPath, filetype);
        }

        public void RunSoftware(string softwarepath)
        {
            //ProcessStartInfo startInfo = new ProcessStartInfo
            //{
            //    FileName = softwarepath,
            //    UseShellExecute = false,
            //    Arguments = @"D:\RefFit\"
            //};
            //Process.Start(startInfo);
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.Start();
            process.StandardInput.WriteLine("cd D:\\RefFit");
            process.StandardInput.WriteLine("run_RefFitTool.exe");
            process.Close();
        }

        public void MoveFile(string sourceDirectory, string targetDirectory, string type)
        {
            string[] datFiles = Directory.GetFiles(sourceDirectory, type);
            foreach (string datFile in datFiles)
            {
                string fileName = Path.GetFileName(datFile);
                string targetPath = Path.Combine(targetDirectory, fileName);
                File.Move(datFile, targetPath);
            }
        }

        public bool CheckCSV(string csvfile1, string csvfile2)
        {
            bool state = false;
            while (true)
            {
                if (File.Exists(csvfile1) && File.Exists(csvfile2))
                {
                    state = true;
                    break;
                }
                Thread.Sleep(1000);
            }
            return state;
        }
    }

    class ExcelHandler
    {
        public string waferID { get; set; }

        private Dictionary<string, int> GetTupleExtremum(List<List<Tuple<string, string, double, double>>> lists)
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

        private void SetChartStyle(ExcelChart chart, Tuple<int, int, int, int> position, List<List<Tuple<string, string, double, double>>> lists)
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

        private void FieldLabel(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1"].Value = "tag";
            worksheet.Cells["B1"].Value = "filename";
            worksheet.Cells["C1"].Value = "wl";
            worksheet.Cells["D1"].Value = "amp";
        }

        private void WhiteCells(ExcelWorksheet worksheet, List<List<Tuple<string, string, double, double>>> lists, int tag_count, int cell_init, int list_group)
        {
            for (int cell_index = 0; cell_index < tag_count; cell_index++)
            {
                worksheet.Cells["A" + (cell_init + cell_index + 1).ToString()].Value = lists[list_group][cell_index].Item1;
                worksheet.Cells["B" + (cell_init + cell_index + 1).ToString()].Value = lists[list_group][cell_index].Item2;
                worksheet.Cells["C" + (cell_init + cell_index + 1).ToString()].Value = lists[list_group][cell_index].Item3;
                worksheet.Cells["D" + (cell_init + cell_index + 1).ToString()].Value = lists[list_group][cell_index].Item4;
            }
        }

        private ExcelRange GetRange(ExcelWorksheet worksheet, string field, int start, int end)
        {
            return worksheet.Cells[field + (start).ToString() + ":" + field + (end).ToString()];
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
                            Tuple<string, string, double, double> rowData = new Tuple<string, string, double, double>(fields[0], fields[1], Convert.ToDouble(fields[2]), Convert.ToDouble(fields[3]));
                            currentChunk.Add(rowData);
                            dataListChunks.Add(currentChunk);
                            tmp = fields[0];
                            currentChunk = new List<Tuple<string, string, double, double>>();
                        }
                    }
                    dataListChunks.Add(currentChunk);
                }
                #region For debug
                //foreach (var chunk in dataListChunks)
                //{
                //    foreach (var tuple in chunk)
                //    {
                //        Console.WriteLine($"({tuple.Item1}, {tuple.Item2}, {tuple.Item3}, {tuple.Item4})");
                //    }
                //}
                #endregion
            }
            return dataListChunks;
        }

        // The name of chart can't be repeated.(ScatterPlot)
        public void WaveToScatterChart(string csvfilepath, string xlsxfilepath)
        {
            List<List<Tuple<string, string, double, double>>> lists = CSVToList(csvfilepath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("output_waveform");
                FieldLabel(worksheet);
                int cell_y = 1;
                for (int list_index = 0; list_index < lists.Count; list_index += 2)
                {
                    string chartname = "ScatterPlot" + list_index.ToString();
                    if (!CheckChartName(worksheet, chartname))
                    {
                        ExcelChart chart = worksheet.Drawings.AddChart(chartname, eChartType.XYScatterLinesNoMarkers);
                        SetChartStyle(chart, new Tuple<int, int, int, int>(cell_y, 0, 5, 0), lists);
                        int tag0_count = lists[list_index].Count;
                        int tag1_count = lists[list_index + 1].Count;
                        WhiteCells(worksheet, lists, tag0_count, cell_y, list_index);
                        WhiteCells(worksheet, lists, tag1_count, cell_y + tag0_count, list_index + 1);
                        var measurementA = GetRange(worksheet, "C", cell_y + 1, cell_y + tag0_count + 1);
                        var measurementB = GetRange(worksheet, "D", cell_y + 1, cell_y + tag0_count + 1);
                        var simulationA = GetRange(worksheet, "C", cell_y + tag0_count + 1, cell_y + tag0_count + tag1_count + 1);
                        var simulationB = GetRange(worksheet, "D", cell_y + tag0_count + 1, cell_y + tag0_count + tag1_count + 1);
                        var measurementseries = (ExcelScatterChartSerie)chart.Series.Add(measurementB, measurementA);
                        var simulationseries = (ExcelScatterChartSerie)chart.Series.Add(simulationB, simulationA);
                        measurementseries.Header = "0-量測";
                        simulationseries.Header = "模擬";
                        cell_y += tag0_count + tag1_count;
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
