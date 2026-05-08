using ClosedXML.Excel;
using System.Configuration;
using System.Text;

namespace TableDataConverter
{
    public partial class Form1 : Form
    {
        //
        static public string pPathScript = string.Empty;
        static public string pPathData = string.Empty;

        //
        TableDataLoaderCreater _mtCreater;

        //
        List<FileInfo> _fileInfos;

        //
        StringBuilder _sb;

        public Form1()
        {
            //
            InitializeComponent();

            //
            var path = Directory.GetParent(Directory.GetCurrentDirectory()).FullName;
            path = $"{path}\\{new DirectoryInfo(AppContext.BaseDirectory).Name.Replace("Tables", "")}";
            pPathScript = $"{path}\\Assets\\Scripts\\_Common\\Table";
            pPathData = $"{path}\\Assets\\Table";

            //
            _mtCreater = new TableDataLoaderCreater();
            _sb = new StringBuilder();

            RefreshFileInfos();
        }

        /// <summary>
        /// 
        /// </summary>
        void RefreshFileInfos()
        {
            //
            if (_fileInfos == null)
                _fileInfos = new List<FileInfo>();

            _fileInfos.Clear();

            //
            var dirInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (var item in dirInfo.GetFiles())
            {
                if (item.Name[0] == '_' && item.Name.IndexOf(".xlsx") >= 0)
                {
                    _fileInfos.Add(item);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtn_Refresh(object sender, EventArgs e)
        {
            RefreshFileInfos();

            MessageBox.Show($"갱신");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtn_Confirm(object sender, EventArgs e)
        {
            //
            foreach (var item in _fileInfos)
            {
                //
                var fileName = item.Name.Replace(".xlsx", "");
                var loadData = new XLWorkbook(item.FullName);

                CreateClass(fileName, loadData);
                CreateData(fileName, loadData);
            }

            //
            _mtCreater.Create(_fileInfos);

            //
            MessageBox.Show($"완료");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="workBook"></param>
        void CreateClass(string fileName, XLWorkbook workBook)
        {
            //
#if !DEBUG
            var fs = new FileStream($"{pPathScript}\\{fileName}.cs", FileMode.Create, FileAccess.Write);
            var sw = new StreamWriter(fs);
#endif

            //
            var worksheet = workBook.Worksheet(1);
            var range = worksheet.RangeUsed();

            if (range == null)
                return;

            var tempVariables = new List<KeyValuePair<string, string>>();
            for (int col = 1; col <= range.ColumnCount(); col++)
            {
                var key = worksheet.Cell(3, col).Value.GetText();
                switch (key)
                {
                    case "int":
                    case "long":
                    case "double":
                    case "float":
                    case "string":
                        var temp = new KeyValuePair<string, string>(key, worksheet.Cell(2, col).Value.GetText());
                        tempVariables.Add(temp);
                        break;
                }
            }
            var data = ClassCode(fileName, tempVariables);
#if !DEBUG
            sw.Write(data);
            sw.Close();
            fs.Close();
#endif
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="variables"></param>
        /// <returns></returns>
        string ClassCode(string name, List<KeyValuePair<string, string>> variables)
        {
            var className = name.Replace(".xlsx", "");

            _sb.Clear();
            _sb.Append($"using System;\r\nusing System.IO;\r\nusing System.Collections.Generic;\r\nusing Newtonsoft.Json;\r\n\r\npublic class {className}\r\n{{");
            _sb.Append("\r\n    public class Values");
            _sb.Append("\r\n    {");
            foreach (var item in variables)
            {
                _sb.AppendFormat("\r\n        public {0} {1} {{ private set; get; }}", item.Key, item.Value);
            }
            _sb.Append("\r\n\r\n        [JsonConstructor]");
            _sb.Append("\r\n        public Values(");
            for (int i = 0; i < variables.Count; i++)
            {
                _sb.AppendFormat("{0} {1}", variables[i].Key, variables[i].Value);
                if (i < variables.Count - 1)
                    _sb.Append(",");
            }
            _sb.Append(")");
            _sb.Append("\r\n        {");
            foreach (var item in variables)
            {
                _sb.AppendFormat("\r\n            this.{0} = {1};", item.Value, item.Value);
            }
            _sb.Append("\r\n        }");
            _sb.Append("\r\n    }");

            _sb.Append($"\r\n\r\n    public static {className}.Values GetItem(int key)\r\n");
            _sb.Append("    {\r\n");
            _sb.Append($"        if (Data.TableDataLoader.Instance._dic{className}.ContainsKey(key))\r\n");
            _sb.Append($"            return Data.TableDataLoader.Instance._dic{className}[key];\r\n");
            _sb.Append("        else\r\n");
            _sb.Append("            return null;\r\n");
            _sb.Append("    }\r\n");
            _sb.Append($"\r\n\r\n    public static List<{className}.Values> GetList()\r\n");
            _sb.Append("    {\r\n");
            _sb.Append($"        return Data.TableDataLoader.Instance._list{className};\r\n");
            _sb.Append("    }\r\n");
            _sb.Append("}");

            return _sb.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="workBook"></param>
        void CreateData(string fileName, XLWorkbook workBook)
        {
#if !DEBUG
            //
            var fs = new FileStream($"{pPathData}\\{fileName}.bytes", FileMode.Create, FileAccess.Write);
            var sw = new StreamWriter(fs);
#endif

            //
            var worksheet = workBook.Worksheet(1);
            var range = worksheet.RangeUsed();

            if (range == null)
                return;

            var rowCount = range.LastRow().RowNumber();
            var columnCount = range.LastColumn().ColumnNumber();            

            _sb.Clear();
            _sb.Append("[");
            var tempVariables = new List<string>();
            for (int row = 4; row <= rowCount; row++)
            {
                _sb.Append("{");
                for (int col = 1; col <= columnCount; col++)
                {
                    var cellValue = worksheet.Cell(row, col).Value;

                    _sb.Append(cellValue.IsText ?
                        DataCode(worksheet.Cell(2, col).GetText(), cellValue.GetText()) :
                        DataCode(worksheet.Cell(2, col).GetText(), cellValue.GetNumber()));

                    if (col < columnCount)
                    {
                        _sb.Append(",");
                    }
                }
                _sb.Append("}");

                if (row < rowCount)
                {
                    _sb.Append(",");
                }
            }
            _sb.Append("]");

#if !DEBUG
            sw.Write(_sb);
            sw.Close();
            fs.Close();
#endif
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        string DataCode(string name, string value)
        {
            var temp = string.Format("\"{0}\":\"{1}\"", name, value);

            return temp;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        string DataCode(string name, double value)
        {
            var temp = string.Format("\"{0}\":{1}", name, value);

            return temp;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
