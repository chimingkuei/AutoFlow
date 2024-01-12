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

        public void SimulateLeftMouseClick(IntPtr windowHandle, int x, int y)
        {
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
        }

        public void SimulateRightMouseClick(IntPtr windowHandle, int x, int y)
        {
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero);
        }

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
        

    }

    class ExcelHandler
    {
        public void CreateXlsx(string filepath, string sheetname)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // 新建一個 Excel 檔案
            var excelFile = new ExcelPackage();
            // 在 Excel 檔案中建立一個工作表
            var worksheet = excelFile.Workbook.Worksheets.Add(sheetname);
            // 在工作表中新增一些資料
            // example:
            worksheet.Cells["A1"].Value = "姓名";
            // 儲存 Excel 檔案
            try
            {
                excelFile.SaveAs(new FileInfo(filepath));
            }
            catch
            {
                Console.WriteLine(filepath + " file is opened! Please close that file.");
            }
        }

        public void ReadXlsx(string filepath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // 路徑到Excel文件
            var fileInfo = new FileInfo(filepath);
            // 使用ExcelPackage讀取Excel文件
            using (var package = new ExcelPackage(fileInfo))
            {
                // 取得第一個工作表
                var worksheet = package.Workbook.Worksheets[0];
                // 讀取單元格的值
                // example:
                //var value = worksheet.Cells[1, 1].Value;
            }
        }

        public void ModifyXlsx(string filepath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // 路徑到Excel文件
            var fileInfo = new FileInfo(filepath);
            // 使用ExcelPackage讀取Excel文件
            using (var package = new ExcelPackage(fileInfo))
            {
                // 取得第一個工作表
                var worksheet = package.Workbook.Worksheets[0];
                // 修改單元格的值
                // example:
                //worksheet.Cells[1, 1].Value = "Hello, world!";
                // 保存Excel文件
                try
                {
                    package.Save();
                }
                catch
                {
                    Console.WriteLine(filepath + " file is opened! Please close that file.");
                }
            }
        }

        private List<Tuple<string, string, string, Color>> SetCategoryColor()
        {
            List<Tuple<string, string, string, Color>> categorycolor = new List<Tuple<string, string, string, Color>>
            {
                new Tuple<string, string, string, Color>("M1:M", "N1:N", "Label A", Color.Red),
                new Tuple<string, string, string, Color>("O1:O", "P1:P", "Label B",Color.Green),
                new Tuple<string, string, string, Color>("Q1:Q", "R1:R", "Label C",Color.Blue)
            };
            return categorycolor;
        }

        private void SetChartStyle(ExcelChart chart)
        {
            chart.SetPosition(1, 0, 4, 0); // Set chart position
            chart.SetSize(600, 400); // Set chart size
            chart.Title.Text = "圖表標題";// Set chart title
            chart.Title.Fill.Color = Color.Cyan;// Set color of chart title
            chart.Legend.Position = eLegendPosition.Right;// Set position of legend
            chart.Legend.Fill.Color = Color.LightGray;// Set color of legend
            chart.XAxis.Title.Text = "X Axis Title";
            chart.XAxis.MajorGridlines.Fill.Color = Color.Gray;
            chart.XAxis.MinorGridlines.Fill.Color = Color.LightGray;
            chart.XAxis.MinValue = 0;
            chart.XAxis.MaxValue = 20;
            chart.YAxis.Title.Text = "Y Axis Title";
            chart.YAxis.MajorGridlines.Fill.Color = Color.Gray;
            chart.YAxis.MinorGridlines.Fill.Color = Color.LightGray;
            chart.YAxis.MinValue = 0;
            chart.YAxis.MaxValue = 20;
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

        // The name of chart can't be repeated.(ScatterPlot)
        public void ScatterChart(string filepath, string sheetname, List<List<Tuple<double, double>>> lists)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetname);
                if (!CheckChartName(worksheet, "ScatterPlot"))
                {
                    var chart = worksheet.Drawings.AddChart("ScatterPlot", eChartType.XYScatter);
                    SetChartStyle(chart);
                    // Set X and Y axis data ranges
                    for (int list_index = 0; list_index < lists.Count; list_index++)
                    {
                        for (int i = 0; i < lists[list_index].Count; i++)
                        {
                            worksheet.Cells[SetCategoryColor()[list_index].Item1[0]+(i+1).ToString()].Value = lists[list_index][i].Item1;
                            worksheet.Cells[SetCategoryColor()[list_index].Item2[0] + (i + 1).ToString()].Value = lists[list_index][i].Item2;
                        }
                        var xRange = worksheet.Cells[SetCategoryColor()[list_index].Item1 + lists[list_index].Count];
                        var yRange = worksheet.Cells[SetCategoryColor()[list_index].Item2 + lists[list_index].Count];
                        var series = (ExcelScatterChartSerie)chart.Series.Add(xRange, yRange);
                        series.Header = SetCategoryColor()[list_index].Item3;
                        series.Marker.Fill.Color = SetCategoryColor()[list_index].Item4;
                        series.Marker.Style = eMarkerStyle.Circle;
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

        public void CSVToList(string csvfilepath, Tuple<int, int> index)
        {
            List<List<Tuple<double, double>>> dataListChunks = new List<List<Tuple<double, double>>>();
            int chunkSize = 256;
            // 確保CSV檔案存在
            if (File.Exists(csvfilepath))
            {
                using (StreamReader reader = new StreamReader(csvfilepath))
                {
                    // 跳過標題行（如果有的話）
                    reader.ReadLine();
                    List<Tuple<double, double>> currentChunk = new List<Tuple<double, double>>();
                    // 讀取CSV檔案中的每一行
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] fields = line.Split(',');
                        Tuple<double, double> rowData = new Tuple<double, double> (Convert.ToDouble(fields[index.Item1]), Convert.ToDouble(fields[index.Item2]));
                        currentChunk.Add(rowData);
                        // 如果已經累積了256筆資料，將它添加到主List中，然後重新創建新的List
                        if (currentChunk.Count == chunkSize)
                        {
                            dataListChunks.Add(currentChunk);
                            currentChunk = new List<Tuple<double, double>>();
                        }
                    }
                }
                // 打印第一個List的第一筆數據（可選）
                //if (dataListChunks.Count > 0 && dataListChunks[0].Count > 0)
                //{
                //    Tuple<double, double> firstData = dataListChunks[0][0];
                //    Console.WriteLine($"Column1: {firstData.Item1}, Column2: {firstData.Item2}");
                //}
            }
            else
            {
                Console.WriteLine("CSV檔案不存在");
            }
        }


    }
}
