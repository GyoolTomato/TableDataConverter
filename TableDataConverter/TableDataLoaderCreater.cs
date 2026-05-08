using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableDataConverter
{
    internal class TableDataLoaderCreater
    {
        //
        StringBuilder _sb;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileInfos"></param>
        public void Create(List<FileInfo> fileInfos)
        {
            if (_sb == null)
                _sb = new StringBuilder();

            //
#if !DEBUG
            var fs = new FileStream($"{Form1.pPathScript}\\TableDataLoader.cs", FileMode.Create, FileAccess.Write);
            var sw = new StreamWriter(fs);
#endif

            //
            _sb.Clear();
            _sb.Append("using System;\r\n");
            _sb.Append("using System.Collections.Generic;\r\n");
            _sb.Append("using UnityEngine.AddressableAssets;\r\n");
            _sb.Append("using Newtonsoft.Json;\r\n");
            _sb.Append("\r\n");
            _sb.Append("namespace Data\r\n");
            _sb.Append("{\r\n");
            _sb.Append("    public class TableDataLoader : Singleton<TableDataLoader>\r\n");
            _sb.Append("    {\r\n");
            foreach (var item in fileInfos)
            {
                //
                var fileName = item.Name.Replace(".xlsx", "");
                _sb.Append($"        public Dictionary<int, {fileName}.Values> _dic{fileName} = new Dictionary<int, {fileName}.Values>();\r\n");
                _sb.Append($"        public List<{fileName}.Values> _list{fileName} = new List<{fileName}.Values>();\r\n");
            }
            _sb.Append("\r\n\r\n        public void Init()\r\n");
            _sb.Append("        {\r\n");
            foreach (var item in fileInfos)
            {
                //
                var fileName = item.Name.Replace(".xlsx", "");
                _sb.Append($"            var temp{fileName} = JsonConvert.DeserializeObject<List<{fileName}.Values>>(Manager_Addressable.Instance.GetTable(\"Assets/Table/{fileName}.bytes\").text);\r\n");
                _sb.Append($"            foreach (var item in temp{fileName})\r\n");
                _sb.Append("            {\r\n");
                _sb.Append($"                _list{fileName}.Add(item);\r\n");
                _sb.Append($"                _dic{fileName}.Add(item.key, item);\r\n");
                _sb.Append("            }\r\n");
            }
            _sb.Append("        }\r\n\r\n");
            _sb.Append("    }\r\n");
            _sb.Append("}");

#if !DEBUG
            sw.Write(_sb.ToString());
            sw.Close();
            fs.Close();
#endif
        }
    }
}
