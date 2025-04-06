using System;
using System.Windows;
using ETABSv1; // 確保你已經引用了 ETABS API 的命名空間
using System.Windows.Input;

namespace YourNamespace.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        public ICommand OpenETABSCommand { get; }

        public MainViewModel()
        {
            OpenETABSCommand = new RelayCommand(OpenETABS);
        }

        private void OpenETABS()
        {
            try
            {
                // 創建 ETABS 應用程序對象
                cOAPI etabsObject = null;
                cHelper myHelper = new Helper();
                etabsObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");

                // 打開 ETABS
                etabsObject.ApplicationStart();
                MessageBox.Show("ETABS 已成功打開！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("打開 ETABS 時發生錯誤: " + ex.Message);
            }
        }
    }
}