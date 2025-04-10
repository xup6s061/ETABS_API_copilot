using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using ETABS_API_copilot.Models;

namespace ETABS_API_copilot.Services
{
    public class JsonService
    {
        public List<Building> LoadFromJson(string filePath)
        {
            if (!File.Exists(filePath)) return new List<Building>();

            try
            {
                var json = File.ReadAllText(filePath);
                return JsonConvert.DeserializeObject<List<Building>>(json) ?? new List<Building>();
            }
            catch (Exception ex)
            {
                // 處理反序列化錯誤
                Console.WriteLine($"讀取 JSON 時發生錯誤: {ex.Message}");
                return new List<Building>();
            }
        }

        public void SaveToJson(string filePath, List<Building> buildings)
        {
            try
            {
                var json = JsonConvert.SerializeObject(buildings, Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
            catch (Exception ex)
            {
                // 處理序列化錯誤
                Console.WriteLine($"儲存 JSON 時發生錯誤: {ex.Message}");
            }
        }
    }
}
