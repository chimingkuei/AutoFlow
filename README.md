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
#For ScatterChart function test
```
 var data1 = new List<Tuple<double, double>>
  {
      Tuple.Create(1.0, 2.0),
      Tuple.Create(2.0, 3.0),
      Tuple.Create(3.0, 4.0),
      Tuple.Create(4.0, 5.0),
      // Add more data points as needed
  };
 var data2 = new List<Tuple<double, double>>
 {
     Tuple.Create(5.0, 8.0),
     Tuple.Create(4.0, 5.0),
     Tuple.Create(5.0, 2.0),
     Tuple.Create(1.0, 1.0),
     // Add more data points as needed
 };
 List<List<Tuple<double, double>>> lists = new List<List<Tuple<double, double>>>
 {
     data1,
     data2
 };
 ExcelHandler EH = new ExcelHandler();
 EH.ScatterChart(@"D:\test.xlsx", "test", lists, ChartType.Wave);
```
#CSV to list
```
string csvFilePath = @"D:\TEST.csv";
EH.CSVToList(csvFilePath, new Tuple<int, int>(1, 2));
```