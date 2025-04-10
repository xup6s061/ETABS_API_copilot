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
            var json = File.ReadAllText(filePath);
            return JsonConvert.DeserializeObject<List<Building>>(json) ?? new List<Building>();
        }

        public void SaveToJson(string filePath, List<Building> buildings)
        {
            var json = JsonConvert.SerializeObject(buildings, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }
    }
}
