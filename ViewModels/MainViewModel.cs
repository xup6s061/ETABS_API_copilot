using System;
using System.Collections.Generic;
using System.Linq;
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
        private double _floorHeight = 3.0; // 每層樓的樓高 (單位: 米)
        private int _numFloors = 5; // 樓層數

        private string _xSpanLengthsInput = "5,5,5"; // Default X spans
        private string _ySpanLengthsInput = "5,5,5"; // Default Y spans

        private double[] _xSpanLengths; // Parsed X spans
        private double[] _ySpanLengths; // Parsed Y spans

        public string XSpanLengthsInput
        {
            get => _xSpanLengthsInput;
            set
            {
                _xSpanLengthsInput = value;
                OnPropertyChanged(nameof(XSpanLengthsInput));
            }
        }

        public string YSpanLengthsInput
        {
            get => _ySpanLengthsInput;
            set
            {
                _ySpanLengthsInput = value;
                OnPropertyChanged(nameof(YSpanLengthsInput));
            }
        }

        public double FloorHeight
        {
            get => _floorHeight;
            set { _floorHeight = value; OnPropertyChanged(nameof(FloorHeight)); }
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

                // 嘗試解析 X 和 Y 方向的跨距輸入
                try
                {
                    _xSpanLengths = ParseSpanLengths(_xSpanLengthsInput);
                    _ySpanLengths = ParseSpanLengths(_ySpanLengthsInput);
                }
                catch
                {
                    MessageBox.Show("無法解析跨距輸入，請確保格式正確，例如: 2@5,3");
                    return;
                }

                if (_xSpanLengths == null || _xSpanLengths.Length == 0 || _ySpanLengths == null || _ySpanLengths.Length == 0)
                {
                    MessageBox.Show("請輸入有效的跨距長度！");
                    return;
                }

                cSapModel sapModel = _etabsObject.SapModel;

                for (int floor = 0; floor < _numFloors; floor++)
                {
                    double z = floor * _floorHeight;
                    double x1 = 0; // 起始位置 x = 0

                    for (int xSpan = 0; xSpan <= _xSpanLengths.Length; xSpan++) // 包含最後一列
                    {
                        double x2 = xSpan < _xSpanLengths.Length ? x1 + _xSpanLengths[xSpan] : x1;
                        double y1 = 0; // 起始位置 y = 0

                        for (int ySpan = 0; ySpan <= _ySpanLengths.Length; ySpan++) // 包含最後一列
                        {
                            double y2 = ySpan < _ySpanLengths.Length ? y1 + _ySpanLengths[ySpan] : y1;

                            // 添加梁（僅在樓層大於 0 且不是最後一列時）
                            if (floor > 0 && xSpan < _xSpanLengths.Length)
                            {
                                string beamNameX = "BeamX";
                                sapModel.FrameObj.AddByCoord(x1, y1, z, x2, y1, z, ref beamNameX, "", "Auto");
                            }

                            if (floor > 0 && ySpan < _ySpanLengths.Length)
                            {
                                string beamNameY = "BeamY";
                                sapModel.FrameObj.AddByCoord(x1, y1, z, x1, y2, z, ref beamNameY, "", "Auto");
                            }

                            // 添加柱子
                            if (floor > 0)
                            {
                                double zPrev = (floor - 1) * _floorHeight;
                                string columnName = "Column";
                                sapModel.FrameObj.AddByCoord(x1, y1, zPrev, x1, y1, z, ref columnName, "", "Auto");
                            }

                            y1 = y2; // 移動到下一個 Y 跨
                        }

                        x1 = x2; // 移動到下一個 X 跨
                    }
                }

                MessageBox.Show("三維建築物構架已成功創建！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("創建建築物構架時發生錯誤: " + ex.Message);
            }
        }

        // 解析跨距輸入
        private double[] ParseSpanLengths(string input)
        {
            var parts = input.Split(',');
            var spanList = new List<double>();

            foreach (var part in parts)
            {
                if (part.Contains("@"))
                {
                    var subParts = part.Split('@');
                    if (subParts.Length == 2 &&
                        int.TryParse(subParts[0], out int count) &&
                        double.TryParse(subParts[1], out double value))
                    {
                        spanList.AddRange(Enumerable.Repeat(value, count));
                    }
                    else
                    {
                        throw new FormatException();
                    }
                }
                else
                {
                    if (double.TryParse(part, out double value))
                    {
                        spanList.Add(value);
                    }
                    else
                    {
                        throw new FormatException();
                    }
                }
            }

            return spanList.ToArray();
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
