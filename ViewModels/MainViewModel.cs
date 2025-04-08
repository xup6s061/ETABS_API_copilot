using System;
using System.Windows;
using System.Windows.Input;
using ETABSv1; // 確保你已經引用了 ETABS API 的命名空間

namespace YourNamespace.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private cOAPI _etabsObject;

        // 命令
        public ICommand OpenETABSCommand { get; }
        public ICommand CreateNewETABSFileCommand { get; }
        public ICommand CreateBuildingFrameworkCommand { get; }

        // 屬性
        private double _spanLength = 5.0; // 每跨的跨距 (單位: 米)
        private double _floorHeight = 3.0; // 每層樓的樓高 (單位: 米)
        private int _numSpans = 3; // 跨數
        private int _numFloors = 5; // 樓層數

        public double SpanLength
        {
            get => _spanLength;
            set { _spanLength = value; OnPropertyChanged(nameof(SpanLength)); }
        }

        public double FloorHeight
        {
            get => _floorHeight;
            set { _floorHeight = value; OnPropertyChanged(nameof(FloorHeight)); }
        }

        public int NumSpans
        {
            get => _numSpans;
            set { _numSpans = value; OnPropertyChanged(nameof(NumSpans)); }
        }

        public int NumFloors
        {
            get => _numFloors;
            set { _numFloors = value; OnPropertyChanged(nameof(NumFloors)); }
        }

        // 構造函數
        public MainViewModel()
        {
            OpenETABSCommand = new RelayCommand(OpenETABS);
            CreateNewETABSFileCommand = new RelayCommand(CreateNewETABSFile, CanCreateNewETABSFile);
            CreateBuildingFrameworkCommand = new RelayCommand(CreateBuildingFramework, CanCreateBuildingFramework);
        }

        // 打開 ETABS
        private void OpenETABS()
        {
            try
            {
                cHelper myHelper = new Helper();
                _etabsObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");
                _etabsObject.ApplicationStart();
                MessageBox.Show("ETABS 已成功打開！");
                CommandManager.InvalidateRequerySuggested(); // 通知命令狀態改變
            }
            catch (Exception ex)
            {
                MessageBox.Show("打開 ETABS 時發生錯誤: " + ex.Message);
            }
        }

        // 創建新檔案
        private void CreateNewETABSFile()
        {
            try
            {
                if (!IsETABSRunning())
                {
                    MessageBox.Show("請先打開 ETABS！");
                    return;
                }

                cSapModel sapModel = _etabsObject?.SapModel;
                if (sapModel == null)
                {
                    MessageBox.Show("無法獲取 ETABS 的 SapModel 對象！");
                    return;
                }

                int ret = sapModel.InitializeNewModel(eUnits.kN_m_C);
                if (ret != 0)
                {
                    MessageBox.Show($"初始化新模型失敗，錯誤代碼: {ret}");
                    return;
                }

                ret = sapModel.File.NewBlank();
                if (ret != 0)
                {
                    MessageBox.Show($"創建新檔案失敗，錯誤代碼: {ret}");
                    return;
                }

                MessageBox.Show("ETABS 新檔案已成功建立！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("建立 ETABS 新檔案時發生錯誤: " + ex.Message);
            }
        }

        // 創建建築物構架
        private void CreateBuildingFramework()
        {
            try
            {
                if (!IsETABSRunning())
                {
                    MessageBox.Show("請先打開 ETABS！");
                    return;
                }

                cSapModel sapModel = _etabsObject.SapModel;

                for (int floor = 0; floor < _numFloors; floor++)
                {
                    double z = floor * _floorHeight;
                    for (int span = 0; span <= _numSpans; span++) // Adjusted to include the last column
                    {
                        double x1 = span * _spanLength;

                        // Add beams between spans (skip for the last span)
                        if (span < _numSpans && floor > 0)
                        {
                            double x2 = (span + 1) * _spanLength;
                            string beamName = "Beam";
                            sapModel.FrameObj.AddByCoord(x1, 0, z, x2, 0, z, ref beamName, "", "Auto");
                        }

                        // Add columns for all spans
                        if (floor > 0)
                        {
                            double zPrev = (floor - 1) * _floorHeight;
                            string columnName = "Column";
                            sapModel.FrameObj.AddByCoord(x1, 0, zPrev, x1, 0, z, ref columnName, "", "Auto");
                        }
                    }
                }

                MessageBox.Show("建築物構架已成功創建！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("創建建築物構架時發生錯誤: " + ex.Message);
            }
        }

        // 檢查 ETABS 是否正在運行
        private bool IsETABSRunning()
        {
            try
            {
                if (_etabsObject == null)
                {
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

        // 命令是否可執行
        private bool CanCreateNewETABSFile() => IsETABSRunning();
        private bool CanCreateBuildingFramework() => IsETABSRunning();
    }
}
