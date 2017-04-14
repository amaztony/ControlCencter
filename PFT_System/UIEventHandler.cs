using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Windows;
using System.Data;
using System.Diagnostics;
using OfficeOpenXml.Style;
using System.Windows.Media;
using System.Windows.Controls.Primitives;
using System.IO.Ports;
using System.Text;
using System.Windows.Threading;
using System.Windows.Documents;
using System.Text.RegularExpressions;
using MySql.Data.MySqlClient;

//Name          :       Physical Fitness Test System
//Environment   :       .NET Framework 4.0
//Author        :       Tony G @SUT
//Email         :       gaozt2014@outlook.com
//Date          :       2016.12 ~ 2017.1

namespace PFT_System
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        int order;
        DataTable dt;

        //string FilePath;    //学生名单Excel路径

        //ExcelPackage package;
        //ExcelWorksheet sheet;

        Stopwatch stopwatch = new Stopwatch();

        private SerialPort serialPort = new SerialPort();
        string indata;

        //数据库相关
        //static string hostAddress = "138.128.199.25";
        //static string userName = "sut";
        //static string userPassword = "g17ZGWz5CN2L66gI";
        string hostAddress;
        string userName;
        string userPassword;
        static string databaseName = "china_pft";
        static string tableName = "sut";
        MySqlConnection conn;

        #region 数据库面板
        private void connectSqlButton_Click(object sender, RoutedEventArgs e)
        {
            if (connectSqlButton.Content.ToString() == "连接数据库")
            {
                try
                {
                    //构造数据库连接字符串
                    hostAddress = hostAddressTextBox.Text;
                    userName = userNameTextBox.Text;
                    userPassword = userPasswordPasswordBox.Password;
                    conn = new MySqlConnection("Database='" + databaseName + "';Data Source=" + hostAddress + ";Persist Security Info=yes;UserId=" + userName + ";PWD=" + userPassword + ";");

                    //连接数据库
                    conn.Open();
                    connectSqlButton.Content = "断开连接";

                    hostAddressTextBox.IsEnabled = false;

                    manualRegButton.IsEnabled = true;

                    StatusBar("成功连接到数据库！", "Yellow");

                }
                catch (Exception ex)
                {
                    StatusBar(ex.Message, "Red");
                    return;
                }
            }
            else if (connectSqlButton.Content.ToString() == "断开连接")
            {
                //断开数据库连接
                conn.Close();

                connectSqlButton.Content = "连接数据库";

                hostAddressTextBox.IsEnabled = true;

                manualRegButton.IsEnabled = false;
                confirmButton.IsEnabled = false;

                StatusBar("已断开连接。", "Blue");
            }
        }

        private void updateSqlButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MySqlCommand cmd = conn.CreateCommand();//命令对象（用来封装需要在数据库执行的语句）

                for (int i = 0; i < order; i++)
                {
                    cmd.CommandText = "UPDATE " + tableName + " SET 50米跑=" + dt.Rows[i]["Score"] + " WHERE 学籍号=" + dt.Rows[i]["ID"];
                    cmd.ExecuteNonQuery();
                    //cmd.ExecuteReader();
                }

                StatusBar("成功提交到数据库！", "Yellow");

            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
                return;
            }
        }
        #endregion

        #region Excel面板
        //private void selectExcelButton_Click(object sender, RoutedEventArgs e)
        //{
        //    Microsoft.Win32.OpenFileDialog dialog =
        //        new Microsoft.Win32.OpenFileDialog();
        //    dialog.Title = "选择包含所有学生名单的Excel";
        //    dialog.Multiselect = false;
        //    dialog.Filter = "Excel 工作薄|*.xlsx";
        //    if (dialog.ShowDialog() == true)
        //    {
        //        FilePath = dialog.FileName;
        //    }
        //    else return;

        //    //File.Copy(FilePath, "连接前备份" + DateTime.Now.ToString("HHmm") + @".bak");
        //    fileName.Text = Path.GetFileName(FilePath);

        //    connectExcelButton.IsEnabled = true;
        //    mergeToCurrentButton.IsEnabled = true;

        //    StatusBar("已选择 " + FilePath, "Blue");
        //}

        //private void connectExcelButton_Click(object sender, RoutedEventArgs e)
        //{
        //    if (connectExcelButton.Content.ToString() == "连接")
        //    {
        //        //FilePath = @"Model.xlsx";

        //        //UI阻塞
        //        //StatusBar("正在连接" + FilePath + "……", "Yellow");
        //        try
        //        {
        //            FileInfo existingFile = new FileInfo(FilePath);
        //            package = new ExcelPackage(existingFile);
        //            sheet = package.Workbook.Worksheets[1];
        //            connectExcelButton.Content = "关闭";
        //            manualRegButton.IsEnabled = true;

        //            StatusBar("成功连接到 " + FilePath, "Yellow");

        //        }
        //        catch (Exception ex)
        //        {
        //            StatusBar(ex.Message, "Red");
        //            return;
        //        }
        //        //catch
        //        //{
        //        //    StatusBar("连接失败，请检查文件是否正确！", "Red");
        //        //}
        //    }
        //    else if (connectExcelButton.Content.ToString() == "关闭")
        //    {
        //        package.Dispose();
        //        connectExcelButton.Content = "连接";
        //        manualRegButton.IsEnabled = false;
        //        confirmButton.IsEnabled = false;

        //        StatusBar("当前文件 " + FilePath, "Blue");
        //    }
        //}

        //private void mergeToCurrentButton_Click(object sender, RoutedEventArgs e)
        //{
        //    MergeToExcel(FilePath);
        //}

        //private void mergeToExcelButton_Click(object sender, RoutedEventArgs e)
        //{
        //    Microsoft.Win32.OpenFileDialog dialog =
        //        new Microsoft.Win32.OpenFileDialog();
        //    dialog.Title = "合并成绩到现有Excel";
        //    dialog.Multiselect = false;
        //    dialog.Filter = "Excel 工作薄|*.xlsx"; // Filter files by extension

        //    string path = string.Empty;

        //    // Process save file dialog box results
        //    if (dialog.ShowDialog() == true)
        //    {
        //        // Save document
        //        path = dialog.FileName;
        //    }
        //    else return;

        //    MergeToExcel(path);
        //}

        private void saveAsReportButton_Click(object sender, RoutedEventArgs e)
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
                File.Copy(path, "报告前备份" + DateTime.Now.ToString("HHmm") + @".bak");
                reportFile.Delete();  // ensures we create a new workbook
                reportFile = new FileInfo(path);
                //File.Create(path);
            }

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(reportFile))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);
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

                    using (var range = worksheet.Cells[1, 1, order + 1, 4])
                    {
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    // save our new workbook and we are done!
                    excelPackage.Save();
                }

                StatusBar("保存成功 " + path, "Yellow");
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
                return;
            }
            //catch
            //{
            //    StatusBar("保存失败，请检查文件是否被其他进程占用！", "Red");
            //}
        }
        #endregion

        #region 操作面板
        private void manualRegButton_Click(object sender, RoutedEventArgs e)
        {
            string ID = manualIDTextBox.Text;
            Register(ID);
        }

        private void confirmButton_Click(object sender, RoutedEventArgs e)
        {
            //若已存在不再重复添加
            if (dt.Select("ID=" + studentIDTextBox.Text).Length.Equals(1))
                return;
            else
            {
                order++;
                runwayTextBlock.Text = order.ToString();
                DataRow dr = dt.NewRow();
                dr["ID"] = studentIDTextBox.Text;
                dr["Name"] = nameTextBox.Text;
                dt.Rows.Add(dr);
            }
        }

        private void queryButton_Click(object sender, RoutedEventArgs e)
        {
            string sheetName = studentIDTextBox.Text;

            Microsoft.Win32.SaveFileDialog dialog =
                new Microsoft.Win32.SaveFileDialog();
            dialog.Title = "体测成绩详单另存为";
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

            FileInfo reportFile = new FileInfo(path);

            if (reportFile.Exists)
            {
                File.Copy(path, "详单生成前备份" + DateTime.Now.ToString("HHmm") + @".bak");
                reportFile.Delete();  // ensures we create a new workbook
                reportFile = new FileInfo(path);
                //File.Create(path);
            }

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(reportFile))
                {
                    string ID = studentIDTextBox.Text;

                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                    //Add the headers
                    worksheet.Cells[1, 1].Value = "学号";
                    worksheet.Cells[1, 2].Value = ID;
                    worksheet.Cells[1, 3].Value = "姓名";
                    worksheet.Cells[1, 4].Value = nameTextBox.Text;
                    worksheet.Cells[1, 5].Value = "专业班级";
                    worksheet.Cells[1, 6].Value = classTextBox.Text;

                    worksheet.Cells[2, 1].Value = "性别代码";//"性别"：1-男，2-女
                    worksheet.Cells[2, 3].Value = "出生日期";
                    worksheet.Cells[2, 5].Value = "家庭住址";

                    worksheet.Cells[3, 1].Value = "身高";
                    worksheet.Cells[3, 3].Value = "体重";
                    worksheet.Cells[3, 5].Value = "肺活量";

                    worksheet.Cells[4, 1].Value = "50米跑";
                    worksheet.Cells[4, 3].Value = "立定跳远";
                    worksheet.Cells[4, 5].Value = "坐位体前屈";

                    worksheet.Cells[5, 1].Value = "800米跑";
                    worksheet.Cells[5, 3].Value = "1000米跑";
                    worksheet.Cells[5, 5].Value = "1分钟仰卧起坐";

                    MySqlCommand cmd = conn.CreateCommand();//命令对象（用来封装需要在数据库执行的语句）
                    cmd.CommandText = "SELECT * FROM " + tableName + " WHERE 学籍号=" + ID;
                    MySqlDataReader sdr = cmd.ExecuteReader();
                    if (sdr.HasRows)
                    {
                        //循环读取返回的数据
                        while (sdr.Read())
                        {
                            try { worksheet.Cells[2, 2].Value = sdr.GetString(sdr.GetOrdinal("性别")); } catch (Exception) { }
                            try { worksheet.Cells[2, 4].Value = sdr.GetString(sdr.GetOrdinal("出生日期")); } catch (Exception) { }
                            try { worksheet.Cells[2, 6].Value = sdr.GetString(sdr.GetOrdinal("家庭住址")); } catch (Exception) { }

                            try { worksheet.Cells[3, 2].Value = sdr.GetString(sdr.GetOrdinal("身高")); } catch (Exception) { }
                            try { worksheet.Cells[3, 4].Value = sdr.GetString(sdr.GetOrdinal("体重")); } catch (Exception) { }
                            try { worksheet.Cells[3, 6].Value = sdr.GetString(sdr.GetOrdinal("肺活量")); } catch (Exception) { }

                            try { worksheet.Cells[4, 2].Value = sdr.GetString(sdr.GetOrdinal("50米跑")); } catch (Exception) { }
                            try { worksheet.Cells[4, 4].Value = sdr.GetString(sdr.GetOrdinal("立定跳远")); } catch (Exception) { }
                            try { worksheet.Cells[4, 6].Value = sdr.GetString(sdr.GetOrdinal("坐位体前屈")); } catch (Exception) { }

                            try { worksheet.Cells[5, 2].Value = sdr.GetString(sdr.GetOrdinal("800米跑")); } catch (Exception) { }
                            try { worksheet.Cells[5, 4].Value = sdr.GetString(sdr.GetOrdinal("1000米跑")); } catch (Exception) { }
                            try { worksheet.Cells[5, 6].Value = sdr.GetString(sdr.GetOrdinal("1分钟仰卧起坐")); } catch (Exception) { }
                        }
                    }
                    sdr.Close();

                    worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells

                    using (var range = worksheet.Cells[1, 1, 5, 6])
                    {
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    excelPackage.Save();
                }

                StatusBar("保存成功 " + path, "Yellow");
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
                return;
            }
            //catch
            //{
            //    StatusBar("保存失败，请检查文件是否被其他进程占用！", "Red");
            //}
        }

        private void assignButton_Click(object sender, RoutedEventArgs e)
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
            //Array.Sort(runways);
            //ArrivComboBox.ItemsSource = runways;
            //ArrivComboBox.SelectedIndex = 0;
        }

        private void stopwatchResetButton_Click(object sender, RoutedEventArgs e)
        {
            stopwatch.Stop();
            stopwatch.Reset();
            stopwatchStartPauseButton.Content = "开始";
            stopwatchStartPauseButton.IsEnabled = true;
        }

        private void stopwatchStartPauseButton_Click(object sender, RoutedEventArgs e)
        {
            if (stopwatchStartPauseButton.Content.ToString() == "开始" || stopwatchStartPauseButton.Content.ToString() == "继续")
            {
                dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 0, 1); //UI更新间隔1ms
                stopwatch.Start();
                stopwatchStartPauseButton.Content = "暂停";
            }
            else if (stopwatchStartPauseButton.Content.ToString() == "暂停")
            {
                dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 1); //UI更新间隔1s
                stopwatch.Stop();
                stopwatchStartPauseButton.Content = "继续";
            }
        }

        private void newMatchButton_Click(object sender, RoutedEventArgs e)
        {
            //复位秒表
            stopwatchResetButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            //清空数据接收
            recvDataRichTextBox.Document.Blocks.Clear();
            //清空数据表格及order变量
            InitializeMemory();
        }
        #endregion

        #region 通信面板
        private void openClosePortButton_Click(object sender, RoutedEventArgs e)
        {
            if (serialPort.IsOpen)
            {
                if (ClosePort())
                {
                    openClosePortButton.Content = "打开";
                }
            }
            else
            {
                if (OpenPort())
                {
                    openClosePortButton.Content = "关闭";
                }
            }
        }

        private void findPortButton_Click(object sender, RoutedEventArgs e)
        {
            FindPorts();
        }

        private void clearRecvDataButton_Click(object sender, RoutedEventArgs e)
        {
            recvDataRichTextBox.Document.Blocks.Clear();
        }

        private void sendTestButton_Click(object sender, RoutedEventArgs e)
        {
            string textToSend = sendDataTextBox.Text;
            if (string.IsNullOrEmpty(textToSend))
            {
                StatusBar("要发送的内容不能为空！", "Red");
                return;
            }
            else
            {
                DataProcess(textToSend);
                ShowMessage(textToSend);
            }
        }

        private void sendDataButton_Click(object sender, RoutedEventArgs e)
        {
            string textToSend = sendDataTextBox.Text;
            if (string.IsNullOrEmpty(textToSend))
            {
                StatusBar("要发送的内容不能为空！", "Red");
                return;
            }
            else SerialPortWrite(textToSend);
        }
        #endregion

        #region 菜单栏
        private void exitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void operationViewMenuItem_Click(object sender, RoutedEventArgs e)
        {
            bool state = operationViewMenuItem.IsChecked;

            if (state == false)
            {
                operationPanel.Visibility = Visibility.Visible;
                statusInfoLength += 40;
                StatusBar("已展开操作面板。", "Blue");
            }
            else
            {
                operationPanel.Visibility = Visibility.Collapsed;
                statusInfoLength -= 40;
                StatusBar("已收起操作面板。", "Blue");
            }

            operationViewMenuItem.IsChecked = !state;
        }

        private void communicationViewMenuItem_Click(object sender, RoutedEventArgs e)
        {
            bool state = communicationViewMenuItem.IsChecked;

            if (state == false)
            {
                communicationPanel.Visibility = Visibility.Visible;
                statusInfoLength += 40;
                StatusBar("已展开通信面板。", "Blue");
            }
            else
            {
                communicationPanel.Visibility = Visibility.Collapsed;
                statusInfoLength -= 40;
                StatusBar("已收起操作面板。", "Blue");
            }

            communicationViewMenuItem.IsChecked = !state;
        }

        private void helpMenuItem_Click(object sender, RoutedEventArgs e)
        {
            //TO-DO: Help
        }

        private void aboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            About about = new About();
            about.ShowDialog();
        }
        #endregion

        #region 状态栏
        /// <summary>
        /// 更新时间信息
        /// </summary>
        private void UpdateTimeDate()
        {
            string timeDateString = "";
            DateTime now = DateTime.Now;
            timeDateString = string.Format("{0}年{1}月{2}日 {3}:{4}:{5}",
                now.Year,
                now.Month.ToString("00"),
                now.Day.ToString("00"),
                now.Hour.ToString("00"),
                now.Minute.ToString("00"),
                now.Second.ToString("00"));

            timeDateTextBlock.Text = timeDateString;
        }

        int statusInfoLength = 100;
        /// <summary>
        /// 信息提示
        /// </summary>
        /// <param name="message">提示信息</param>
        private void StatusBar(string message)
        {
            statusInfoTextBlock.Text = GetSubString(message, statusInfoLength);
        }

        /// <summary>
        /// 信息提示 区分颜色模式
        /// </summary>
        /// <param name="message">提示信息</param>
        private void StatusBar(string message, string mode)
        {
            if (mode == "Blue")
            {
                // #FF007ACC
                statusBar.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0x7A, 0xCC));
            }
            else if (mode == "Yellow")
            {
                // #FFCA5100
                statusBar.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0xCA, 0x51, 0x00));
            }
            else if (mode == "Red")
            {
                // #FF68217A
                statusBar.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0xFF, 0x21, 0x2A));
            }
            StatusBar(message);
        }

        public static string GetSubString(string origStr, int endIndex)
        {
            if (origStr == null || origStr.Length == 0 || endIndex < 0)
                return "";
            int bytesCount = Encoding.GetEncoding("gb2312").GetByteCount(origStr);
            if (bytesCount > endIndex)
            {
                int readyLength = 0;
                int byteLength;
                for (int i = 0; i < origStr.Length; i++)
                {
                    byteLength = Encoding.GetEncoding("gb2312").GetByteCount(new char[] { origStr[i] });
                    readyLength += byteLength;
                    if (readyLength == endIndex)
                    {
                        origStr = origStr.Substring(0, i + 1) + "...";
                        break;
                    }
                    else if (readyLength > endIndex)
                    {
                        origStr = origStr.Substring(0, i) + "...";
                        break;
                    }
                }
            }
            return origStr;
        }
        #endregion

        #region 其他
        public void InitializeMemory()
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

        //private void MergeToExcel(string path)
        //{
        //    //合并前备份目标文件，防止错误覆盖
        //    File.Copy(path, "合并前备份" + DateTime.Now.ToString("HHmm") + @".bak");

        //    FileInfo outFile = new FileInfo(path);

        //    //FileInfo outFile = new FileInfo(@"out.xlsx");
        //    //if (!outFile.Exists)    //不合适
        //    //{
        //    //    File.Copy(FilePath, @"out.xlsx");
        //    //}
        //    try
        //    {
        //        using (ExcelPackage excelPackage = new ExcelPackage(outFile))
        //        {
        //            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];
        //            for (int i = 0; i < order; i++)
        //            {
        //                var query1 = (from cell in worksheet.Cells["d:d"] where cell.Value.Equals(dt.Rows[i]["ID"]) select cell);
        //                foreach (var cell in query1)
        //                {
        //                    int RowIDx = int.Parse(cell.Address.Substring(1));  //取得行号
        //                    worksheet.Cells[RowIDx, 13].Value = dt.Rows[i]["Score"];
        //                    break;
        //                }
        //            }
        //            excelPackage.Save();
        //        }

        //        StatusBar("合并成功 " + path, "Yellow");
        //    }
        //    catch (Exception ex)
        //    {
        //        StatusBar(ex.Message, "Red");
        //        return;
        //    }
        //    //catch
        //    //{
        //    //    StatusBar("合并失败，请检查文件是否正确！", "Red");

        //    //}
        //}

        private void Register(string ID)
        {
            //try-catch
            MySqlCommand cmd = conn.CreateCommand();//命令对象（用来封装需要在数据库执行的语句）
            cmd.CommandText = "SELECT * FROM " + tableName + " WHERE 学籍号=" + ID;
            try
            {
                MySqlDataReader sdr = cmd.ExecuteReader();
                if (sdr.HasRows)
                {
                    //循环读取返回的数据
                    while (sdr.Read())
                    {
                        studentIDTextBox.Text = ID;
                        nameTextBox.Text = sdr.GetString(sdr.GetOrdinal("姓名"));
                        classTextBox.Text = sdr.GetString(sdr.GetOrdinal("班级名称"));
                        sexComboBox.ItemsSource = new string[] { "男", "女" };
                        sexComboBox.SelectedIndex = sdr.GetInt32(sdr.GetOrdinal("性别")) - 1;
                    }
                }
                sdr.Close();

                confirmButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
            }

            //var query1 = (from cell in sheet.Cells["d:d"] where cell.Value.Equals(ID) select cell);
            //foreach (var cell in query1)
            //{
            //    int RowIDx = int.Parse(cell.Address.Substring(1));  //取得行号
            //    studentIDTextBox.Text = ID;
            //    nameTextBox.Text = sheet.Cells[RowIDx, 6].Value.ToString();
            //    classTextBox.Text = sheet.Cells[RowIDx, 3].Value.ToString();
            //    sexComboBox.ItemsSource = new string[] { "男", "女" };
            //    sexComboBox.SelectedIndex = sheet.Cells[RowIDx, 7].GetValue<Int16>() - 1;
            //    break;
            //}
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
        #endregion

        #region 串口相关
        private void InitSerialPort()
        {
            serialPort.DataReceived += SerialPort_DataReceived;
            FindPorts();
        }

        private void ConfigurePort()
        {
            serialPort.PortName = portsComboBox.Text;
            serialPort.BaudRate = 9600;
            serialPort.Parity = Parity.None;
            serialPort.DataBits = 8;
            serialPort.StopBits = StopBits.One;
            serialPort.Encoding = Encoding.Default;
        }

        private bool OpenPort()
        {
            bool flag = false;
            ConfigurePort();

            try
            {
                serialPort.Open();
                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                StatusBar(string.Format("成功打开端口{0}。", serialPort.PortName), "Yellow");
                flag = true;
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
            }

            return flag;
        }

        private bool ClosePort()
        {
            bool flag = false;

            try
            {
                serialPort.Close();
                StatusBar(string.Format("成功关闭端口{0}。", serialPort.PortName), "Yellow");
                flag = true;
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
            }

            return flag;
        }

        private void FindPorts()
        {
            portsComboBox.ItemsSource = SerialPort.GetPortNames();
            if (portsComboBox.Items.Count > 0)
            {
                portsComboBox.SelectedIndex = 0;
                portsComboBox.IsEnabled = true;

                StatusBar(string.Format("查找到可以使用的端口{0}个。", portsComboBox.Items.Count.ToString()), "Blue");
            }
            else
            {
                portsComboBox.IsEnabled = false;
                StatusBar("没有查找到可用端口，请刷新。", "Red");
            }
        }

        private bool SerialPortWrite(string textData)
        {
            if (serialPort == null)
            {
                return false;
            }

            if (serialPort.IsOpen == false)
            {
                StatusBar("串口未打开，无法发送数据。", "Red");
                return false;
            }

            try
            {
                //serialPort.DiscardOutBuffer();
                //serialPort.DiscardInBuffer();

                serialPort.Write(textData);

                // 报告发送成功的消息，提示用户。
                StatusBar(string.Format("成功发送：{0}。", textData), "Yellow");
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
                return false;
            }

            return true;
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = sender as SerialPort;
            indata = sp.ReadExisting();

            //委托UI线程显示
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new DelegateShowMessage(DelegateEvents));
        }

        private delegate void DelegateShowMessage();  //定义委托

        private void DelegateEvents()
        {
            DataProcess(indata);
            ShowMessage(indata);
        }

        private void ShowMessage(string data)    //要让UI线程完成的事情
        {
            Paragraph p = new Paragraph();
            Run r = new Run(data);
            p.Inlines.Add(r);
            recvDataRichTextBox.Document.Blocks.Add(p);

            //recvDataRichTextBox.AppendText(data + "\r\n");
        }

        private void DataProcess(string data)
        {
            if (Regex.IsMatch(data, @"^ID\d{9}$"))
            {
                string ID = data.Substring(2);
                StatusBar("学号为 " + ID + " 的学生进行检录。", "Blue");
                try
                {
                    Register(ID);
                }
                catch (Exception ex)
                {
                    StatusBar(ex.Message, "Red");
                }
            }
            else if (Regex.IsMatch(data, @"^START$"))
            {
                SerialPortWrite("Started!");    //回复开始信号
                StatusBar("开始计时！", "Blue");
                stopwatchResetButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                stopwatchStartPauseButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            }
            else if (Regex.IsMatch(data, @"^AR\d{1,2}$"))
            {
                string arrive = data.Substring(2);

                //若已存在不再重复添加
                if (dt.Select("Runway=" + arrive + " and Score is not null").Length.Equals(1))
                    return;
                else
                {
                    StatusBar("跑道号为 " + arrive + " 的学生已经冲线。", "Yellow");

                    try
                    {
                        DataRow dr = dt.Rows[int.Parse(arrive) - 1];
                        dr["Score"] = (stopwatch.ElapsedMilliseconds / 1000.0).ToString("F1");
                    }
                    catch (Exception ex)
                    {
                        StatusBar(ex.Message, "Red");
                    }
                }
            }
            else
            {
                StatusBar("数据格式有误！", "Red");
            }

        }
        #endregion
    }
}