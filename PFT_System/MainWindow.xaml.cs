using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

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
        private DispatcherTimer dispatcherTimer = null;
        public MainWindow()
        {
            InitializeComponent();

            InitializeMemory();

            InitSerialPort();

            dispatcherTimer = new DispatcherTimer();
            dispatcherTimer.Tick += new EventHandler(OnTimedEvent);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 1); //UI更新间隔1s
            dispatcherTimer.Start();
        }

        private void OnTimedEvent(object sender, EventArgs e)
        {
            if (stopwatch.ElapsedMilliseconds >= 10900)
            {
                stopwatch.Stop();
                stopwatchTextBlock.Text = "超时";
                stopwatchStartPauseButton.IsEnabled = false;
            }
            else
            {
                stopwatchTextBlock.Text = stopwatch.ElapsedMilliseconds.ToString("00:00:000").Substring(0, 8);
            }
            UpdateTimeDate();
        }
    }
}
