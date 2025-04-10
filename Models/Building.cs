using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ETABS_API_copilot.Models
{
    public class Building
    {
        public string BuildingName { get; set; }
        public string FloorHeights { get; set; } // 樓層高度 (以逗號分隔)
        public string XSpans { get; set; }      // X方向跨距 (以逗號分隔)
        public string YSpans { get; set; }      // Y方向跨距 (以逗號分隔)
    }
}