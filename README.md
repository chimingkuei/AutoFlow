# AutoFlow
#SetChartStyle
 ```
chart.SetPosition(1, 0, 4, 0); // Set chart position
chart.SetSize(600, 400); // Set chart size
chart.Title.Text = waferID;// Set chart title
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
```
#Generate scatter map
```
string csvpath = "put your csv path"; 
string wavepath = "put your xlsx path";
List<List<Tuple<double, double>>> lists = EH.CSVToList(csvpath, new Tuple<int, int>(2, 3));
EH.WaveToScatterChart(wavepath, "sheetname", lists);
```
```
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
    Do.SimulateRightMouseClick(Convert.ToInt32(TextBoxDispatcherGetValue(Coordinate_X)), Convert.ToInt32(TextBoxDispatcherGetValue(Coordinate_Y)));
    System.Windows.Forms.SendKeys.SendWait("D:\\oCam");
    Do.SimulateLeftMouseClick(899, 156);
    Thread.Sleep(3000);
}
else
{
    Console.WriteLine($"{TextBoxDispatcherGetValue(Window_Name)} Window can't be found.");
}
```