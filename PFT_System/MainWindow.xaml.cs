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
        public MainWindow()
        {
            InitializeComponent();

            InitializeMemory();

            InitSerialPort();
        }
    }
}
