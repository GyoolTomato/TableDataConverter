using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableDataConverter
{
    internal class ManagerTableCreater
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
            var fs = new FileStream($"{Form1.pPathScript}\\Manager_Table.cs", FileMode.Create, FileAccess.Write);
            var sw = new StreamWriter(fs);

            //
            _sb.Clear();
            _sb.Append("using System;\r\n");
            _sb.Append("using System.Collections.Generic;\r\n");
            _sb.Append("using UnityEngine.AddressableAssets;\r\n");
            _sb.Append("using Newtonsoft.Json;\r\n");
            _sb.Append("\r\n");
            _sb.Append("public class Manager_Table : Singleton<Manager_Table>\r\n");
            _sb.Append("{\r\n");
            foreach (var item in fileInfos)
            {
                //
                var fileName = item.Name.Replace(".xlsx", "");
                _sb.Append($"    public Dictionary<int, {fileName}> _dic{fileName} = new Dictionary<int, {fileName}>();\r\n");
            }
            _sb.Append("\r\n\r\n    public void Init()\r\n");
            _sb.Append("    {\r\n");
            foreach (var item in fileInfos)
            {
                //
                var fileName = item.Name.Replace(".xlsx", "");
                _sb.Append($"        var temp = JsonConvert.DeserializeObject<List<{fileName}>>(Manager_Addressable.Instance.GetTable(\"{fileName}\").text);\r\n");
                _sb.Append($"        foreach (var item in temp)\r\n");
                _sb.Append("        {\r\n");
                _sb.Append($"            _dic{fileName}.Add(item.pKey, item);\r\n");
                _sb.Append("        }\r\n");
            }
            _sb.Append("    }\r\n\r\n");
            _sb.Append("}");


            sw.Write(_sb.ToString());
            sw.Close();
            fs.Close();
        }
    }
}
