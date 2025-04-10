using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ETABS_API_copilot.Models
{
    public class Material
    {
        public string BuildingName { get; set; }
        public string MaterialName { get; set; }
        public double Density { get; set; }
        public double ElasticModulus { get; set; }
        public double Strength { get; set; }
        public double PoissonRatio { get; set; } // Added property
        public double CoefficientThermalExpansion { get; set; } // Added property
    }

    public class SectionProperty
    {
        public string BuildingName { get; set; } // 所屬建築物名稱
        public string SectionName { get; set; } // 斷面名稱
        public string Material { get; set; } // 此斷面所使用之材料名稱
        public double Width { get; set; } // 寬度
        public double Height { get; set; } // 高度
    }

    public class Building
    {
        public string BuildingName { get; set; }
        public string FloorHeights { get; set; } // 樓層高度 (以逗號分隔)
        public string XSpans { get; set; }      // X方向跨距 (以逗號分隔)
        public string YSpans { get; set; }      // Y方向跨距 (以逗號分隔)
        public List<Material> Materials { get; set; } // 材料清單
        public List<SectionProperty> SectionProperties { get; set; } // 斷面性質清單
    }
}