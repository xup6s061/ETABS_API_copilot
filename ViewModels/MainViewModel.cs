using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ETABSv1;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using Microsoft.Win32;
using ETABS_API_copilot.Models;
using ETABS_API_copilot.Services;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace ETABS_API_copilot.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private cOAPI _etabsObject;
        private readonly JsonService _jsonService = new JsonService();
        private Building _selectedBuilding;

        // 命令
        public ICommand OpenETABSCommand { get; }
        public ICommand CreateNewETABSFileCommand { get; }
        public ICommand CreateBuildingFrameworkCommand { get; }
        public ObservableCollection<Building> Buildings { get; set; } = new ObservableCollection<Building>();
        public ObservableCollection<Material> AllMaterials { get; set; } = new ObservableCollection<Material>();
        public ObservableCollection<SectionProperty> AllSectionProperties { get; set; } = new ObservableCollection<SectionProperty>();
        public RelayCommand LoadJsonCommand { get; }
        public RelayCommand SaveJsonCommand { get; }
        public RelayCommand ExecuteSelectedBuildingsCommand { get; }


        // 屬性
        private readonly double[] _xSpanLengths; // Parsed X spans
        private readonly double[] _ySpanLengths; // Parsed Y spans
        private readonly double[] _floorHeights; // Parsed floor heights


        // 構造函數
        public MainViewModel()
        {
            OpenETABSCommand = new RelayCommand(OpenETABS);
            CreateNewETABSFileCommand = new RelayCommand(CreateNewETABSFile, CanCreateNewETABSFile);
            CreateBuildingFrameworkCommand = new RelayCommand(() => CreateBuildingFramework(_xSpanLengths, _ySpanLengths, _floorHeights), CanCreateBuildingFramework);
            LoadJsonCommand = new RelayCommand(LoadJson);
            SaveJsonCommand = new RelayCommand(SaveJson);
            ExecuteSelectedBuildingsCommand = new RelayCommand(ExecuteSelectedBuildings);

        }

        // 打開 ETABS
        private void OpenETABS()
        {
            try
            {
                cHelper myHelper = new Helper();
                _etabsObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");
                _etabsObject.ApplicationStart();
                //System.Windows.MessageBox.Show("ETABS 已成功打開！");
                CommandManager.InvalidateRequerySuggested(); // 通知命令狀態改變
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("打開 ETABS 時發生錯誤: " + ex.Message);
            }
        }
        public Building SelectedBuilding
        {
            get => _selectedBuilding;
            set
            {
                _selectedBuilding = value;
                OnPropertyChanged(nameof(SelectedBuilding));
            }
        }
        // 創建新檔案
        private void CreateNewETABSFile()
        {
            try
            {
                if (!IsETABSRunning())
                {
                    System.Windows.MessageBox.Show("請先打開 ETABS！");
                    return;
                }

                cSapModel sapModel = _etabsObject?.SapModel;
                if (sapModel == null)
                {
                    System.Windows.MessageBox.Show("無法獲取 ETABS 的 SapModel 對象！");
                    return;
                }

                int ret = sapModel.InitializeNewModel(eUnits.kN_m_C);
                if (ret != 0)
                {
                    System.Windows.MessageBox.Show($"初始化新模型失敗，錯誤代碼: {ret}");
                    return;
                }

                ret = sapModel.File.NewBlank();
                if (ret != 0)
                {
                    System.Windows.MessageBox.Show($"創建新檔案失敗，錯誤代碼: {ret}");
                    return;
                }

                //System.Windows.MessageBox.Show("ETABS 新檔案已成功建立！");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("建立 ETABS 新檔案時發生錯誤: " + ex.Message);
            }
        }
        private void LoadJson()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON Files (*.json)|*.json",
                Title = "選擇 JSON 檔案"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var jsonService = new JsonService();
                var buildings = jsonService.LoadFromJson(openFileDialog.FileName);

                if (buildings.Any())
                {
                    Buildings.Clear();
                    AllMaterials.Clear();
                    AllSectionProperties.Clear();

                    foreach (var building in buildings)
                    {
                        Buildings.Add(building);

                        // 將材料加入 AllMaterials，並附加 BuildingName
                        foreach (var material in building.Materials)
                        {
                            material.BuildingName = building.BuildingName; // 動態添加 BuildingName
                            AllMaterials.Add(material);
                        }

                        // 將斷面加入 AllSectionProperties，並附加 BuildingName
                        foreach (var section in building.SectionProperties)
                        {
                            section.BuildingName = building.BuildingName; // 動態添加 BuildingName
                            AllSectionProperties.Add(section);
                        }
                    }
                }
            }
        }

        private void SaveJson()
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "JSON Files (*.json)|*.json",
                Title = "儲存 JSON 檔案"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                // 將 AllMaterials 和 AllSectionProperties 的修改同步回 Buildings
                foreach (var building in Buildings)
                {
                    building.Materials = AllMaterials
                        .Where(m => m.BuildingName == building.BuildingName)
                        .ToList();

                    building.SectionProperties = AllSectionProperties
                        .Where(s => s.BuildingName == building.BuildingName)
                        .ToList();
                }

                // 儲存到 JSON 檔案
                _jsonService.SaveToJson(saveFileDialog.FileName, new List<Building>(Buildings));
            }
        }

        // 創建建築物構架
        private void CreateBuildingFramework(double[] xSpanLengths, double[] ySpanLengths, double[] floorHeights)
        {
            try
            {
                if (!IsETABSRunning())
                {
                    System.Windows.MessageBox.Show("請先打開 ETABS！");
                    return;
                }

                if (xSpanLengths == null || xSpanLengths.Length == 0 ||
                    ySpanLengths == null || ySpanLengths.Length == 0 ||
                    floorHeights == null || floorHeights.Length == 0)
                {
                    System.Windows.MessageBox.Show("請輸入有效的跨距或樓層高度！");
                    return;
                }

                cSapModel sapModel = _etabsObject.SapModel;

                double zPrev = 0; // 初始化 zPrev 為 0

                for (int floor = 0; floor < floorHeights.Length; floor++)
                {
                    double z = zPrev + floorHeights[floor]; // 當前樓層的高度為前一樓層高度加上當前樓層高度
                    double x1 = 0; // 起始位置 x = 0

                    for (int xSpan = 0; xSpan <= xSpanLengths.Length; xSpan++) // 包含最後一列
                    {
                        double x2 = xSpan < xSpanLengths.Length ? x1 + xSpanLengths[xSpan] : x1;
                        double y1 = 0; // 起始位置 y = 0

                        for (int ySpan = 0; ySpan <= ySpanLengths.Length; ySpan++) // 包含最後一列
                        {
                            double y2 = ySpan < ySpanLengths.Length ? y1 + ySpanLengths[ySpan] : y1;

                            // 添加梁（僅在樓層大於等於 0 且不是最後一列時）
                            if (floor >= 0 && xSpan < xSpanLengths.Length)
                            {
                                string beamNameX = "BeamX";
                                sapModel.FrameObj.AddByCoord(x1, y1, z, x2, y1, z, ref beamNameX, "", "Auto");
                            }

                            if (floor >= 0 && ySpan < ySpanLengths.Length)
                            {
                                string beamNameY = "BeamY";
                                sapModel.FrameObj.AddByCoord(x1, y1, z, x1, y2, z, ref beamNameY, "", "Auto");
                            }

                            // 添加柱子
                            string columnName = "Column";
                            sapModel.FrameObj.AddByCoord(x1, y1, zPrev, x1, y1, z, ref columnName, "", "Auto");

                            y1 = y2; // 移動到下一個 Y 跨
                        }

                        x1 = x2; // 移動到下一個 X 跨
                    }

                    zPrev = z; // 更新 zPrev 為當前樓層的高度
                }

                //System.Windows.MessageBox.Show("三維建築物構架已成功創建！");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("創建建築物構架時發生錯誤: " + ex.Message);
            }
        }

        // 解析輸入的跨距
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
        private void ExecuteSelectedBuildings()
        {
            if (Buildings == null || Buildings.Count == 0)
            {
                System.Windows.MessageBox.Show("沒有可執行的建物參數！");
                return;
            }

            var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "選擇存檔路徑"
            };

            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string folderPath = folderDialog.SelectedPath;

                foreach (var building in Buildings)
                {
                    try
                    {
                        // 創建新檔案
                        CreateNewETABSFile();

                        // 解析參數
                        double[] xSpanLengths = ParseSpanLengths(building.XSpans);
                        double[] ySpanLengths = ParseSpanLengths(building.YSpans);
                        double[] floorHeights = ParseSpanLengths(building.FloorHeights);

                        // 呼叫 CreateBuildingFramework，傳入解析後的參數
                        CreateBuildingFramework(xSpanLengths, ySpanLengths, floorHeights);

                        // 建立材料
                        foreach (var material in AllMaterials.Where(m => m.BuildingName == building.BuildingName))
                        {
                            CreateMaterial(material);
                        }

                        // 建立斷面屬性
                        foreach (var section in AllSectionProperties.Where(s => s.BuildingName == building.BuildingName))
                        {
                            CreateSectionProperty(section);
                        }

                        // 儲存檔案
                        string filePath = Path.Combine(folderPath, $"{building.BuildingName}.edb");
                        int ret = _etabsObject.SapModel.File.Save(filePath);

                        if (ret != 0)
                        {
                            System.Windows.MessageBox.Show($"儲存檔案失敗: {building.BuildingName}，錯誤代碼: {ret}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"執行建物 {building.BuildingName} 時發生錯誤: {ex.Message}");
                    }
                }

                System.Windows.MessageBox.Show("所有建物已成功執行並儲存！");
            }
        }
        // 建立材料
        private void CreateMaterial(Material material)
        {
            try
            {
                var sapModel = _etabsObject.SapModel;

                // 根據材料類型設定 eMatType
                eMatType materialType;
                switch (material.MaterialType)
                {
                    case "Steel":
                        materialType = eMatType.Steel;
                        break;
                    case "Concrete":
                        materialType = eMatType.Concrete;
                        break;
                    case "Aluminum":
                        materialType = eMatType.Aluminum;
                        break;
                    case "ColdFormed":
                        materialType = eMatType.ColdFormed;
                        break;
                    case "Rebar":
                        materialType = eMatType.Rebar;
                        break;
                    case "Tendon":
                        materialType = eMatType.Tendon;
                        break;
                    case "Masonry":
                        materialType = eMatType.Masonry;
                        break;
                    case "Other":
                        materialType = eMatType.NoDesign; // 假設 "Other" 對應 NoDesign
                        break;
                    default:
                        throw new ArgumentException($"未知的材料類型: {material.MaterialType}");
                }

                // 設定材料屬性
                int ret = sapModel.PropMaterial.SetMaterial(material.MaterialName, materialType);
                if (ret == 0)
                {
                    sapModel.PropMaterial.SetMPIsotropic(material.MaterialName, material.ElasticModulus * 98, material.PoissonRatio, material.CoefficientThermalExpansion);
                    sapModel.PropMaterial.SetWeightAndMass(material.MaterialName, 0, material.Density);

                    // 僅當材料類型為 Concrete 時，設定 SetOConcrete
                    if (materialType == eMatType.Concrete)
                    {
                        sapModel.PropMaterial.SetOConcrete(material.MaterialName, material.Strength * 98, false, 0, 1, 2, 0.002, 0.003);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"建立材料 {material.MaterialName} 時發生錯誤: {ex.Message}");
            }
        }
        // 建立斷面屬性
        private void CreateSectionProperty(SectionProperty section)
        {
            try
            {
                var sapModel = _etabsObject.SapModel;
                int ret = sapModel.PropFrame.SetRectangle(section.SectionName, section.Material, section.Height, section.Width);
                if (ret != 0)
                {
                    System.Windows.MessageBox.Show($"建立斷面屬性 {section.SectionName} 失敗，錯誤代碼: {ret}");
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"建立斷面屬性 {section.SectionName} 時發生錯誤: {ex.Message}");
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
