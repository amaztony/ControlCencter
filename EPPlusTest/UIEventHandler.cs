using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Windows;
using System.Data;
using System.Diagnostics;
using OfficeOpenXml.Style;

namespace EPPlusTest
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        int order;
        DataTable dt;
        public void NewMatch()
        {
            order = 0;
            if (dt != null) dt.Dispose();
            dt = new DataTable();

            dt.Columns.Add("ID");
            dt.Columns.Add("Name");
            dt.Columns.Add("Runway");
            dt.Columns.Add("Score");
            mainDataGrid.ItemsSource = dt.DefaultView;
        }

        string FilePath;
        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog =
                new Microsoft.Win32.OpenFileDialog();
            dialog.Title = "选择包含所有学生名单的Excel";
            dialog.Multiselect = false;
            dialog.Filter = "Excel 工作薄|*.xlsx";
            if (dialog.ShowDialog() == true)
            {
                FilePath = dialog.FileName;
            }
            else return;

            //File.Copy(FilePath, "连接前备份" + DateTime.Now.ToString("HHmm") + @".bak");
        }

        //应手动释放
        ExcelPackage excelPackage;
        ExcelWorksheet sheet;
        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            if (ConnectButton.Content.ToString() == "关闭")
            {
                excelPackage.Dispose();
                ConnectButton.Content = "连接";
            }
            else if (ConnectButton.Content.ToString() == "连接")
            {
                //FilePath = @"Model.xlsx";
                FileInfo existingFile = new FileInfo(FilePath);
                excelPackage = new ExcelPackage(existingFile);
                sheet = excelPackage.Workbook.Worksheets[1];
                ConnectButton.Content = "关闭";
            }
        }

        private void SimuRegButton_Click(object sender, RoutedEventArgs e)
        {
            //IDTextBox.Text = ID;

            string ID = IDTextBox.Text;

            //string FilePath = @"Model.xlsx";
            //FileInfo existingFile = new FileInfo(FilePath);
            //using (ExcelPackage package = new ExcelPackage(existingFile))
            //{
                //ExcelWorksheet sheet = excelPackage.Workbook.Worksheets[1];

                var query1 = (from cell in sheet.Cells["d:d"] where cell.Value.Equals(ID) select cell);

                foreach (var cell in query1)
                {
                    int RowIDx = int.Parse(cell.Address.Substring(1));  //取得行号
                    NameTextBox.Text = sheet.Cells[RowIDx, 6].Value.ToString();
                    ClassTextBox.Text = sheet.Cells[RowIDx, 3].Value.ToString();
                    SexComboBox.ItemsSource = new string[] { "男", "女" };
                    SexComboBox.SelectedIndex = sheet.Cells[RowIDx, 7].GetValue<Int16>() - 1;
                    break;
                }
            //}
        }
        
        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            order++;
            RunwayTextBox.Text = order.ToString();
            DataRow dr = dt.NewRow();
            dr["ID"] = IDTextBox.Text;
            dr["Name"] = NameTextBox.Text;
            dt.Rows.Add(dr);
        }

        private void AssignButton_Click(object sender, RoutedEventArgs e)
        {
            int[] runways = BuildRandomSequence(1, order);
            for (int i = 0; i < order; i++)
            {
                DataRow dr = dt.Rows[i];
                dr["Runway"] = runways[i];
            }

            //按跑道号排序
            DataView dv = dt.DefaultView;
            dv.Sort = "Runway";
            dt = dv.ToTable();
            mainDataGrid.ItemsSource = dt.DefaultView;  //重新绑定

            //模拟到达组合框 添加序号
            Array.Sort(runways);
            ArrivComboBox.ItemsSource = runways;
            ArrivComboBox.SelectedIndex = 0;
        }

        Stopwatch stopwatch = new Stopwatch();
        private void SimuStartButton_Click(object sender, RoutedEventArgs e)
        {
            //stopwatch.Stop();
            stopwatch.Reset();
            stopwatch.Start();
        }

        private void SimuArrivButton_Click(object sender, RoutedEventArgs e)
        {
            //arrive应从通信端获取
            int arrive = ArrivComboBox.SelectedIndex;

            DataRow dr = dt.Rows[arrive];

            dr["Score"] = (stopwatch.ElapsedMilliseconds / 1000.0).ToString("F1");
        }

        private void RestartButton_Click(object sender, RoutedEventArgs e)
        {
            stopwatch.Stop();
            NewMatch();
        }

        private int[] BuildRandomSequence(int low, int high)
        {
            int x = 0, tmp = 0;
            if (low > high)
            {
                tmp = low;
                low = high;
                high = tmp;
            }
            int[] array = new int[high - low + 1];
            for (int i = low; i <= high; i++)
            {
                array[i - low] = i;
            }
            Random rand = new Random();
            for (int i = array.Length - 1; i > 0; i--)
            {
                x = rand.Next(0, i + 1);
                tmp = array[i];
                array[i] = array[x];
                array[x] = tmp;
            }
            return array;
        }

        private void ReportButton_Click(object sender, RoutedEventArgs e)
        {
            string sheetName = DateTime.Now.ToString("yyyyMMdd-HH-mm");

            Microsoft.Win32.SaveFileDialog dialog =
                new Microsoft.Win32.SaveFileDialog();
            dialog.Title = "报告另存为   （警告：选择已有文件将会覆盖）";
            dialog.FileName = sheetName; // Default file name
            dialog.DefaultExt = ".xlsx"; // Default file extension
            dialog.Filter = "Excel 工作薄|*.xlsx"; // Filter files by extension
            
            string path = string.Empty;

            // Process save file dialog box results
            if (dialog.ShowDialog() == true)
            {
                // Save document
                path = dialog.FileName;
            }
            else return;
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
            //    @"\" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

            FileInfo reportFile = new FileInfo(path);

            if (reportFile.Exists)
            {
                reportFile.Delete();  // ensures we create a new workbook
                reportFile = new FileInfo(path);
                //File.Create(path);
            }

            using (ExcelPackage package = new ExcelPackage(reportFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
                //Add the headers
                worksheet.Cells[1, 1].Value = "学号";
                worksheet.Cells[1, 2].Value = "姓名";
                worksheet.Cells[1, 3].Value = "跑道号";
                worksheet.Cells[1, 4].Value = "成绩";

                for (int i = 0; i < order; i++)
                {
                    DataRow dr = dt.Rows[i];
                    worksheet.Cells[i + 2, 1].Value = dr["ID"];
                    worksheet.Cells[i + 2, 2].Value = dr["Name"];
                    worksheet.Cells[i + 2, 3].Value = dr["Runway"];
                    worksheet.Cells[i + 2, 4].Value = dr["Score"];
                }

                worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells
                
                // save our new workbook and we are done!
                package.Save();
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog =
                new Microsoft.Win32.OpenFileDialog();
            dialog.Title = "合并成绩到现有Excel";
            dialog.Multiselect = false;
            dialog.Filter = "Excel 工作薄|*.xlsx"; // Filter files by extension

            string path = string.Empty;

            // Process save file dialog box results
            if (dialog.ShowDialog() == true)
            {
                // Save document
                path = dialog.FileName;
            }
            else return;

            //合并前备份目标文件，防止错误覆盖
            File.Copy(path, "合并前备份" + DateTime.Now.ToString("HHmm") + @".bak");

            FileInfo outFile = new FileInfo(path);

            //FileInfo outFile = new FileInfo(@"out.xlsx");
            //if (!outFile.Exists)    //不合适
            //{
            //    File.Copy(FilePath, @"out.xlsx");
            //}

            using (ExcelPackage package = new ExcelPackage(outFile))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                for (int i = 0; i < order; i++)
                {
                    var query1 = (from cell in sheet.Cells["d:d"] where cell.Value.Equals(dt.Rows[i]["ID"]) select cell);
                    foreach (var cell in query1)
                    {
                        int RowIDx = int.Parse(cell.Address.Substring(1));  //取得行号
                        sheet.Cells[RowIDx, 13].Value = dt.Rows[i]["Score"];
                        break;
                    }
                }
                package.Save();
            }
        }

    }
}