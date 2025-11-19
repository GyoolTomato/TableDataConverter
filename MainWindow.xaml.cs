using System.IO;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TableDataConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //
        static public string pPathScript = string.Empty;
        static public string pPathData = string.Empty;

        //
        TableLoaderCreater _tlCreater;

        //
        List<FileInfo> _fileInfos;

        //
        StringBuilder _sb;


        /// <summary>
        /// 
        /// </summary>
        public MainWindow()
        {
            //
            InitializeComponent();

            //
            var path = Directory.GetParent(Directory.GetCurrentDirectory()).FullName;
            path = $"{path}\\{new DirectoryInfo(AppContext.BaseDirectory).Name.Replace("Tables", "")}";
            pPathScript = $"{path}\\Assets\\Scripts\\_Common\\Table";
            pPathData = $"{path}\\Assets\\Table";

            //
            _tlCreater = new TableLoaderCreater();
            _sb = new StringBuilder();

            Refresh();
        }

        /// <summary>
        /// 
        /// </summary>
        void Refresh()
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
        void OnBtn_Refresh(object sender, RoutedEventArgs e)
        {
            Refresh();

            MessageBox.Show($"갱신");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void OnBtn_Confirm(object sender, RoutedEventArgs e)
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
            _tlCreater.Create(_fileInfos);

            //
            MessageBox.Show($"클릭");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="workBook"></param>
        void CreateClass(string fileName, XLWorkbook workBook)
        {
            //
            var fs = new FileStream($"{pPathScript}\\{fileName}.cs", FileMode.Create, FileAccess.Write);
            var sw = new StreamWriter(fs);

            //
            var worksheet = workBook.Worksheet(1);
            var range = worksheet.RangeUsed();

            if (range == null)
                return;

            var tempVariables = new List<string>();
            for (int col = 1; col <= range.ColumnCount(); col++)
            {
                tempVariables.Add(ClassVariable(worksheet.Cell(3, col).Value.GetText(), worksheet.Cell(2, col).Value.GetText()));
            }

            sw.Write(ClassCode(fileName, tempVariables));
            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="variables"></param>
        /// <returns></returns>
        string ClassCode(string name, List<string> variables)
        {
            var className = name.Replace(".xlsx", "");

            _sb.Clear();
            _sb.Append($"using System;\r\nusing System.IO;\r\n\r\npublic class {className}\r\n{{");

            foreach (var item in variables)
            {
                _sb.Append(item);
            }

            _sb.Append($"\r\n\r\n    public static {className} GetItem(int key)\r\n");
            _sb.Append("    {\r\n");
            _sb.Append($"        if (Manager_Table.Instance._dic{className}.ContainsKey(key))\r\n");
            _sb.Append($"            return Manager_Table.Instance._dic{className}[key];\r\n");
            _sb.Append("        else\r\n");
            _sb.Append("            return null;\r\n");
            _sb.Append("    }\r\n");
            _sb.Append("}");

            return _sb.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        string ClassVariable(string type, string name)
        {
            //
            var init = string.Empty;
            switch (type)
            {
                case "int":
                case "long":
                case "double":
                    init = "0";
                    break;
                case "float":
                    init = "0f";
                    break;
                default:
                    return string.Empty;
            }

            //
            _sb.Clear();
            _sb.AppendFormat("\r\n    public {0} {1} {{ set; get; }} = {2};", type, name, init);

            return _sb.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="workBook"></param>
        void CreateData(string fileName, XLWorkbook workBook)
        {
            //
            var fs = new FileStream($"{pPathData}\\{fileName}.bytes", FileMode.Create, FileAccess.Write);
            var sw = new StreamWriter(fs);

            //
            var worksheet = workBook.Worksheet(1);
            var range = worksheet.RangeUsed();

            if (range == null)
                return;

            _sb.Clear();
            _sb.Append("[");
            var tempVariables = new List<string>();
            for (int row = 4; row <= range.RowCount(); row++)
            {
                _sb.Append("{");
                for (int col = 1; col <= range.ColumnCount(); col++)
                {
                    var cellValue = worksheet.Cell(row, col).Value;
                    _sb.Append(cellValue.IsText ? 
                        DataCode(worksheet.Cell(2, col).GetText(), cellValue.GetText()) :
                        DataCode(worksheet.Cell(2, col).GetText(), cellValue.GetNumber()));

                    if (col < range.ColumnCount())
                    {
                        _sb.Append(",");
                    }                    
                }
                _sb.Append("}");

                if (row < range.RowCount())
                {
                    _sb.Append(",");
                }
            }
            _sb.Append("]");

            

            sw.Write(_sb);
            sw.Close();
            fs.Close();
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
    }
}