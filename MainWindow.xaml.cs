using System;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Windows;
using ETABS_API_copilot.ViewModels; // 引用包含 MainViewModel 的命名空間

namespace ETABS_API_copilot
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel(); // 設置 DataContext
        }
    }
}
