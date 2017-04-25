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
using System.Windows.Controls;

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
        static string pftTableName = "sut_pft";
        static string cardInfoTableName = "card_info";
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

                foreach (DataRow dr in dt.Rows)
                {
                    for (int i = 3; i < dt.Columns.Count; i++)  //0：机器号；1：学号；2：姓名
                    {
                        if (dr[i] != null)
                        {
                            string itemName = string.Empty;
                            switch (dt.Columns[i].ColumnName)
                            {
                                case "Height":
                                    itemName = "身高";
                                    break;
                                case "Weight":
                                    itemName = "体重";
                                    break;
                                case "Vital":
                                    itemName = "肺活量";
                                    break;
                                case "Run800":
                                    itemName = "800米跑";
                                    break;
                                case "Run1000":
                                    itemName = "1000米跑";
                                    break;
                                case "Run50":
                                    itemName = "50米跑";
                                    break;
                                case "Jump":
                                    itemName = "立定跳远";
                                    break;
                                case "Flexion":
                                    itemName = "坐位体前屈";
                                    break;
                                case "SitUps":
                                    itemName = "一分钟仰卧起坐";
                                    break;
                                case "PullUp":
                                    itemName = "引体向上";
                                    break;
                            }
                            cmd.CommandText = "UPDATE " + pftTableName + " SET " + itemName + "=" + dr[i] + " WHERE 学籍号=" + dt.Rows[i]["ID"];
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                //for (int i = 0; i < order; i++)
                //{
                //    cmd.CommandText = "UPDATE " + pftTableName + " SET 50米跑=" + dt.Rows[i]["Run50"] + " WHERE 学籍号=" + dt.Rows[i]["ID"];
                //    cmd.ExecuteNonQuery();
                //    //cmd.ExecuteReader();
                //}

                StatusBar("成功提交到数据库！", "Yellow");

            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
                return;
            }
        }

        private void exportEduButton_Click(object sender, RoutedEventArgs e)
        {
            string sheetName = "沈阳工业大学体测结果";

            Microsoft.Win32.SaveFileDialog dialog =
                new Microsoft.Win32.SaveFileDialog();
            dialog.Title = "上报模板另存为";
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
                File.Copy(path, "上报模板导出前备份" + DateTime.Now.ToString("HHmm") + @".bak");
                reportFile.Delete();  // ensures we create a new workbook
                //File.Create(path);
            }
            try
            {
                File.Copy("exportModel.xlsx", path);
                reportFile = new FileInfo(path);
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
            }

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(reportFile))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];

                    MySqlCommand cmd = conn.CreateCommand();//命令对象（用来封装需要在数据库执行的语句）
                    cmd.CommandText = "SELECT * FROM " + pftTableName;
                    MySqlDataReader sdr = cmd.ExecuteReader();
                    int excelRow = 2;
                    if (sdr.HasRows)
                    {
                        //循环读取返回的数据
                        while (sdr.Read())
                        {
                            for (int excelColumn = 1; excelColumn < 20; excelColumn++)
                            {
                                try { worksheet.Cells[excelRow, excelColumn].Value = sdr.GetString(excelColumn - 1); } catch (Exception) { }
                            }

                            excelRow++;
                        }
                    }
                    sdr.Close();

                    excelPackage.Save();
                }

                StatusBar("保存成功 " + path, "Yellow");
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
                return;
            }
        }

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
                //File.Create(path);
            }
            try
            {
                File.Copy("reportModel.xlsx", path);
                reportFile = new FileInfo(path);
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
            }

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(reportFile))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];

                    worksheet.Cells[1, 2].Value = userNameTextBox.Text; //负责人
                    worksheet.Cells[2, 2].Value = DateTime.Now.ToString("yyyy/MM/dd");    //日期

                    for (int i = 0; i < order; i++)
                    {
                        for (int j = 1; j < dt.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 5, j].Value = dt.Rows[i][j];
                        }
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
            try
            {
                string manualID = manualIDTextBox.Text;
                int manualMachineNumber = int.Parse(manualMachineNumberTextBox.Text);
                Register(manualID, manualMachineNumber);
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
            }
        }

        private void confirmButton_Click(object sender, RoutedEventArgs e)
        {
            //若已存在ID，变更机器
            if (dt.Select("ID=" + studentIDTextBox.Text).Length.Equals(1))
            {
                DataRow[] arrayDR = dt.Select("ID=" + studentIDTextBox.Text);
                foreach (DataRow dr in arrayDR)
                {
                    dr["Machine"] = machineNumberTextBox.Text;
                    break;
                }
            }
            else if (dt.Select("Machine=" + machineNumberTextBox.Text).Length.Equals(1))
            {
                StatusBar("机器仍在使用中！", "Red");
            }
            else
            {
                order++;
                DataRow dr = dt.NewRow();
                dr["Machine"] = machineNumberTextBox.Text;
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
                //File.Create(path);
            }
            try
            {
                File.Copy("queryModel.xlsx", path);
                reportFile = new FileInfo(path);
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
            }

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(reportFile))
                {
                    string ID = studentIDTextBox.Text;

                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];
                    
                    worksheet.Cells[4, 2].Value = ID;
                    worksheet.Cells[4, 4].Value = nameTextBox.Text;
                    worksheet.Cells[5, 2].Value = classTextBox.Text;

                    MySqlCommand cmd = conn.CreateCommand();//命令对象（用来封装需要在数据库执行的语句）
                    cmd.CommandText = "SELECT * FROM " + pftTableName + " WHERE 学籍号=" + ID;
                    MySqlDataReader sdr = cmd.ExecuteReader();
                    if (sdr.HasRows)
                    {
                        //循环读取返回的数据
                        while (sdr.Read())
                        {
                            try { worksheet.Cells[5, 4].Value = sdr.GetString(sdr.GetOrdinal("性别")); } catch (Exception) { }
                            try { worksheet.Cells[6, 2].Value = sdr.GetString(sdr.GetOrdinal("出生日期")); } catch (Exception) { }
                            try { worksheet.Cells[6, 4].Value = sdr.GetString(sdr.GetOrdinal("家庭住址")); } catch (Exception) { }

                            try { worksheet.Cells[8, 2].Value = sdr.GetString(sdr.GetOrdinal("身高")); } catch (Exception) { }
                            try { worksheet.Cells[9, 2].Value = sdr.GetString(sdr.GetOrdinal("体重")); } catch (Exception) { }
                            try { worksheet.Cells[10, 2].Value = sdr.GetString(sdr.GetOrdinal("50米跑")); } catch (Exception) { }
                            try { worksheet.Cells[11, 2].Value = sdr.GetString(sdr.GetOrdinal("800米跑")); } catch (Exception) { }
                            try { worksheet.Cells[12, 2].Value = sdr.GetString(sdr.GetOrdinal("1000米跑")); } catch (Exception) { }

                            try { worksheet.Cells[8, 4].Value = sdr.GetString(sdr.GetOrdinal("肺活量")); } catch (Exception) { }
                            try { worksheet.Cells[9, 4].Value = sdr.GetString(sdr.GetOrdinal("立定跳远")); } catch (Exception) { }
                            try { worksheet.Cells[10, 4].Value = sdr.GetString(sdr.GetOrdinal("坐位体前屈")); } catch (Exception) { }
                            try { worksheet.Cells[11, 4].Value = sdr.GetString(sdr.GetOrdinal("1分钟仰卧起坐")); } catch (Exception) { }
                            try { worksheet.Cells[12, 4].Value = sdr.GetString(sdr.GetOrdinal("引体向上")); } catch (Exception) { }
                        }
                    }
                    sdr.Close();
                    
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

        private void rightViewMenuItem_Click(object sender, RoutedEventArgs e)
        {
            bool state = rightViewMenuItem.IsChecked;

            if (state == false)
            {
                operationPanel.Visibility = Visibility.Visible;
                communicationPanel.Visibility = Visibility.Visible;

                statusInfoLength += 40;
                StatusBar("已展开右侧面板。", "Blue");
            }
            else
            {
                operationPanel.Visibility = Visibility.Collapsed;
                communicationPanel.Visibility = Visibility.Collapsed;

                statusInfoLength -= 40;
                StatusBar("已收起右侧面板。", "Blue");
            }

            rightViewMenuItem.IsChecked = !state;
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

        int statusInfoLength = 120;
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

            dt.Columns.Add("Machine");  //Hidden
            dt.Columns.Add("ID");
            dt.Columns.Add("Name");
            dt.Columns.Add("Height");
            dt.Columns.Add("Weight");
            dt.Columns.Add("Vital");
            dt.Columns.Add("Run800");
            dt.Columns.Add("Run1000");
            dt.Columns.Add("Run50");
            dt.Columns.Add("Jump");
            dt.Columns.Add("Flexion");
            dt.Columns.Add("SitUps");
            dt.Columns.Add("PullUp");
            mainDataGrid.ItemsSource = dt.DefaultView;
        }

        private void Register(string studentID, int machineNumber)
        {
            //try-catch
            MySqlCommand cmd = conn.CreateCommand();//命令对象（用来封装需要在数据库执行的语句）
            cmd.CommandText = "SELECT * FROM " + pftTableName + " WHERE 学籍号=" + studentID;
            try
            {
                MySqlDataReader sdr = cmd.ExecuteReader();
                if (sdr.HasRows)
                {
                    //循环读取返回的数据
                    while (sdr.Read())
                    {
                        studentIDTextBox.Text = studentID;
                        nameTextBox.Text = sdr.GetString(sdr.GetOrdinal("姓名"));
                        classTextBox.Text = sdr.GetString(sdr.GetOrdinal("班级名称"));
                        if (sdr.GetInt32(sdr.GetOrdinal("性别")) == 1)
                        {
                            sexTextBox.Text = "男";
                        }
                        else if (sdr.GetInt32(sdr.GetOrdinal("性别")) == 2)
                        {
                            sexTextBox.Text = "女";
                        }
                        machineNumberTextBox.Text = machineNumber.ToString();
                    }
                }
                sdr.Close();

                confirmButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                StatusBar(ex.Message, "Red");
            }
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

        //Test string:"M01I00D2502689101E"
        private string queryStudentID(string cardNumber)
        {
            string studentID = String.Empty;
            MySqlCommand cmd = conn.CreateCommand();//命令对象（用来封装需要在数据库执行的语句）
            cmd.CommandText = "SELECT * FROM " + cardInfoTableName + " WHERE 卡号=" + cardNumber;
            MySqlDataReader sdr = cmd.ExecuteReader();
            if (sdr.HasRows)
            {
                //循环读取返回的数据
                while (sdr.Read())
                {
                    try { studentID = sdr.GetString(0); } catch (Exception) { }
                }
            }
            sdr.Close();
            return studentID;
        }

        private void DataProcess(string data)
        {
            if (Regex.IsMatch(data, @"^M\d{2}I\d{2}D\d{2,10}E$"))
            {
                data = GetSubString(data, data.Length - 1); //去掉结尾的E
                int machineNumber = int.Parse(data.Substring(1, 2));    ///得到机器号
                int itemNumber = int.Parse(data.Substring(4, 2));   //得到项目代号
                int dataContent = 0;
                if (itemNumber != 0)
                {
                    dataContent = int.Parse(data.Substring(7)); //得到项目数据
                }

                DataRow[] arrayDR;
                switch (itemNumber) //下版本应当不再区分项目编号，直接得到项目名称和项目成绩
                {
                    case 0: //检录：内容为学号，9位
                        string studentID = queryStudentID(data.Substring(7));
                        StatusBar("学号为 " + studentID + " 的学生进行检录。", "Blue");
                        try
                        {
                            Register(studentID, machineNumber);
                        }
                        catch (Exception ex)
                        {
                            StatusBar(ex.Message, "Red");
                        }
                        break;
                    case 1: //身高：4位数，以毫米为单位
                        string height = (dataContent / 10.0).ToString("f1");
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["Height"] = height;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生身高是" + height + "厘米。", "Yellow");
                        break;
                    case 2: //体重：3位数，以百克为单位
                        string weight = (dataContent / 10.0).ToString("f1");
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["Weight"] = weight;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生体重是" + weight + "千克。", "Yellow");
                        break;
                    case 3: //肺活量：4位数
                        string vital = dataContent.ToString();
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["Vital"] = vital;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生肺活量是" + vital + "。", "Yellow");
                        break;
                    case 4: //800米：3位数，以秒为单位
                        TimeSpan ts800 = new TimeSpan(0, 0, dataContent);
                        string run800 = ts800.Minutes + "'" + ts800.Seconds;
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["Run800"] = run800;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生800米成绩是" + run800 + "。", "Yellow");
                        break;
                    case 5: //1000米：3位数，以秒为单位
                        TimeSpan ts1000 = new TimeSpan(0, 0, dataContent);
                        string run1000 = ts1000.Minutes + "'" + ts1000.Seconds;
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["Run1000"] = run1000;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生1000米成绩是" + run1000 + "。", "Yellow");
                        break;
                    case 6: //50米：3位数，以百豪秒为单位
                        string run50 = (dataContent / 10.0).ToString("f1");
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["Run50"] = run50;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生50米成绩是" + run50 + "秒。", "Yellow");
                        break;
                    case 7: //立定跳远：3位数，以厘米为单位
                        string jump = dataContent.ToString("f2");
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["Jump"] = jump;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生立定跳远成绩是" + jump + "厘米。", "Yellow");
                        break;
                    case 8: //坐位体前屈：3位数，以毫米为单位
                        string flexion = (dataContent / 10.0).ToString("f1");
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["Flexion"] = flexion;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生坐位体前屈成绩是" + flexion + "厘米。", "Yellow");
                        break;
                    case 9: //仰卧起坐：个
                        string sitUps = dataContent.ToString();
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["SitUps"] = sitUps;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生仰卧起坐成绩是" + sitUps + "个。", "Yellow");
                        break;
                    case 10: //引体向上：个
                        string pullUp = dataContent.ToString();
                        arrayDR = dt.Select("Machine=" + machineNumber);
                        foreach (DataRow dr in arrayDR)
                        {
                            dr["PullUp"] = pullUp;
                            break;
                        }
                        StatusBar("机器号为 " + machineNumber + " 的学生引体向上成绩是" + pullUp + "个。", "Yellow");
                        break;
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