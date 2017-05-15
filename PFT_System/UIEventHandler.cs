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
        string FilePath = @"card_info.xlsx";    //学生名单Excel路径

        private SerialPort serialPort = new SerialPort();
        string indata;

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

        #region 状态栏
        /// <summary>
        /// 信息提示
        /// </summary>
        /// <param name="message">提示信息</param>
        private void StatusBar(string message)
        {
            // #FF007ACC
            statusBar.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0x7A, 0xCC));
            statusInfoTextBlock.Text = message;
        }

        /// <summary>
        /// 信息提示 三种
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
            statusInfoTextBlock.Text = message;
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
            serialPort.BaudRate = 4800;
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

                StatusBar(string.Format("查找到可以使用的端口{0}个。", portsComboBox.Items.Count.ToString()));
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

        private void queryStudentID(uint cardNumber, out string stuID, out string stuName)
        {
            stuID = String.Empty;
            stuName = String.Empty;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                var query1 = (from cell in sheet.Cells["d:d"] where (cell.Text).Equals(cardNumber.ToString()) select cell);
                foreach (var cell in query1)
                {
                    int RowIDx = int.Parse(cell.Address.Substring(1));  //取得行号
                    stuID = sheet.Cells[RowIDx, 1].Text;
                    stuName = sheet.Cells[RowIDx, 2].Text;
                    break;
                }
            }
        }

        private void DataProcess(string data)
        {
            //UID检录匹配，格式为1A2B3C4D
            if ((Regex.IsMatch(data, @"^M\d{2}I00D[0-9a-fA-F]{8}E$")))
            {
                int machineNumber = int.Parse(data.Substring(1, 2));    ///得到机器号
                data = data.Substring(7, 8);    //截取UID
                data = data.Substring(6, 2) + data.Substring(4, 2) + data.Substring(2, 2) + data.Substring(0, 2);   //倒序
                uint dataContent = 0;
                dataContent = Convert.ToUInt32(data, 16);  //得到倒序UID转的数字
                
                try
                {
                    string studentID = String.Empty;
                    string studentName = String.Empty;
                    queryStudentID(dataContent, out studentID, out studentName);
                    StatusBar("学号为 " + studentID + " 的学生开始检录，其姓名为 " + studentName + "。", "Blue");
                }
                catch (Exception ex)
                {
                    StatusBar(ex.Message, "Red");
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