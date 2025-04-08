using System;
using System.Windows;
using ETABSv1; // 確保你已經引用了 ETABS API 的命名空間
using System.Windows.Input;

namespace YourNamespace.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private cOAPI _etabsObject;
        public ICommand OpenETABSCommand { get; }
        public ICommand CreateNewETABSFileCommand { get; }

        public MainViewModel()
        {
            OpenETABSCommand = new RelayCommand(OpenETABS);
            CreateNewETABSFileCommand = new RelayCommand(CreateNewETABSFile, CanCreateNewETABSFile);
        }

        private void OpenETABS()
        {
            try
            {
                // 創建 ETABS 應用程序對象
                cHelper myHelper = new Helper();
                _etabsObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");

                // 打開 ETABS
                _etabsObject.ApplicationStart();
                MessageBox.Show("ETABS 已成功打開！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("打開 ETABS 時發生錯誤: " + ex.Message);
            }
        }

        private bool CanCreateNewETABSFile()
        {
            return IsETABSRunning();
        }

        private void CreateNewETABSFile()
        {
            try
            {
                if (!IsETABSRunning())
                {
                    MessageBox.Show("請先打開 ETABS！");
                    return;
                }

                // 創建新檔案
                cSapModel sapModel = _etabsObject.SapModel;
                int ret = sapModel.InitializeNewModel(eUnits.kN_m_C);
                ret = sapModel.File.NewBlank();

                MessageBox.Show("ETABS 新檔案已成功建立！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立 ETABS 新檔案時發生錯誤: " + ex.Message);
            }
        }

        private bool IsETABSRunning()
        {
            try
            {
                if (_etabsObject == null)
                {
                    // 嘗試連接到已經運行的 ETABS 實例
                    cHelper myHelper = new Helper();
                    _etabsObject = myHelper.GetObject("CSI.ETABS.API.ETABSObject");
                }

                return _etabsObject != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
