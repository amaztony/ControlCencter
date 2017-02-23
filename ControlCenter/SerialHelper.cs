using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace ControlCenter
{
    public partial class MainWindow : Window
    {
        private void InitSerialPort()
        {
            serialPort.DataReceived += SerialPort_DataReceived;
            FindPorts();
        }

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

        private void refreshPortButton_Click(object sender, RoutedEventArgs e)
        {
            FindPorts();
        }

        private void sendDataButton_Click(object sender, RoutedEventArgs e)
        {
            string textToSend = sendDataTextBox.Text;
            if (string.IsNullOrEmpty(textToSend))
            {
                Alert("要发送的内容不能为空！");
                //return false;
            }
            else SerialPortWrite(textToSend);
        }

        private void clearRecvDataBoxButton_Click(object sender, RoutedEventArgs e)
        {
            recvDataTextBox.Clear();
        }

        private void ConfigurePort()
        {
            serialPort.PortName = GetSelectedPortName();
            serialPort.BaudRate = 9600;
            serialPort.Parity = Parity.None;
            serialPort.DataBits = 8;
            serialPort.StopBits = StopBits.One;
            serialPort.Encoding = Encoding.Default;
        }

        private string GetSelectedPortName()
        {
            return portsComboBox.Text;
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
                Information(string.Format("成功打开端口{0}, 波特率{1}。", serialPort.PortName, serialPort.BaudRate.ToString()));
                flag = true;
            }
            catch (Exception ex)
            {
                Alert(ex.Message);
            }

            return flag;
        }

        private bool ClosePort()
        {
            bool flag = false;

            try
            {
                serialPort.Close();
                Information(string.Format("成功关闭端口{0}。", serialPort.PortName));
                flag = true;
            }
            catch (Exception ex)
            {
                Alert(ex.Message);
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
                Information(string.Format("查找到可以使用的端口{0}个。", portsComboBox.Items.Count.ToString()));
            }
            else
            {
                portsComboBox.IsEnabled = false;
                Alert("没有查找到可用端口，请刷新。");
            }
        }

        private SerialPort serialPort = new SerialPort();
        private bool SerialPortWrite(string textData)
        {
            if (serialPort == null)
            {
                return false;
            }

            if (serialPort.IsOpen == false)
            {
                Alert("串口未打开，无法发送数据。");
                return false;
            }

            try
            {
                //serialPort.DiscardOutBuffer();
                //serialPort.DiscardInBuffer();

                serialPort.Write(textData);

                // 报告发送成功的消息，提示用户。
                Information(string.Format("成功发送：{0}。", textData));
            }
            catch (Exception ex)
            {
                Alert(ex.Message);
                return false;
            }

            return true;
        }

        #region 串口接收事件
        string indata;
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = sender as SerialPort;
            indata = sp.ReadExisting();

            //委托UI线程显示
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new DelegateShowMessage(ShowMessage));
        }

        private delegate void DelegateShowMessage();  //定义委托
        private void ShowMessage()    //要让UI线程完成的事情
        {
            recvDataTextBox.Text = indata;
        }

        #endregion

        #region 状态栏消息显示
        //常规消息显示
        private void Information(string message)
        {
            if (serialPort.IsOpen)
            {
                // #FFCA5100
                statusBar.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0xCA, 0x51, 0x00));
            }
            else
            {
                // #FF007ACC
                statusBar.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0x00, 0x7A, 0xCC));
            }
            statusInfoTextBlock.Text = message;
        }

        //警告消息显示
        private void Alert(string message)
        {
            // #FF68217A
            statusBar.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0xFF, 0x21, 0x2A));
            statusInfoTextBlock.Text = message;
        }
        #endregion
    }
}
