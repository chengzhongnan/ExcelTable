using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelTable
{
    /// <summary>
    /// 用於存放從 AppConfig.xml 讀取的設定。
    /// </summary>
    public class AppSettings
    {
        public string InputPath { get; set; }
        public string OutputPath { get; set; }
        public string AssemblyName { get; set; }
    }

    /// <summary>
    /// 負責從 XML 檔案載入應用程式設定。
    /// </summary>
    public static class ConfigLoader
    {
        public static AppSettings Load(string filePath)
        {
            var doc = XDocument.Load(filePath);
            var config = doc.Element("Configuration");

            var settings = new AppSettings
            {
                InputPath = config.Element("Paths").Element("InputPath").Value,
                OutputPath = config.Element("Paths").Element("OutputPath").Value,
                AssemblyName = config.Element("Settings").Element("AssemblyName").Value
            };

            // 將可能存在的相對路徑轉換為絕對路徑，使後續操作更可靠
            settings.InputPath = Path.GetFullPath(settings.InputPath);
            settings.OutputPath = Path.GetFullPath(settings.OutputPath);

            return settings;
        }

        public static void CreateDefault(string filePath)
        {
            var doc = new XDocument(
                new XComment(" Application Configuration File "),
                new XElement("Configuration",
                    new XElement("Paths",
                        new XComment(" Excel files input directory. Can be a relative or absolute path. "),
                        new XElement("InputPath", "xlsx"),
                        new XComment(" Directory for all generated output. "),
                        new XElement("OutputPath", "output")
                    ),
                    new XElement("Settings",
                        new XComment(" The name of the generated assembly and its root namespace. "),
                        new XElement("AssemblyName", "ExcelTable.Configs")
                    )
                )
            );
            doc.Save(filePath);
        }
    }
}
