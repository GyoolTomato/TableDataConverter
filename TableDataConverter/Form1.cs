using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Configuration;
using System.Globalization;
using System.Reflection.Emit;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace TableDataConverter
{
    public partial class Form1 : Form
    {
        //
        static public string pPathGlobalData = string.Empty;
        static public string pPathScript = string.Empty;
        static public string pPathTableData = string.Empty;
        static public string pPathTableLocalData = string.Empty;

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
            pPathGlobalData = $"{path}\\Assets\\Scripts\\_Common\\GlobalData";
            pPathScript = $"{path}\\Assets\\Scripts\\_Common\\Tables";
            pPathTableData = $"{path}\\Assets\\Tables";
            pPathTableLocalData = $"{path}\\Assets\\Resources\\Tables";

            // UI에서는 파일 종류별로 출력 폴더를 하나씩 선택한다.
            pPathGlobalData = pPathScript;
            pPathTableLocalData = pPathTableData;
            textBoxScriptPath.Text = pPathScript;
            textBoxBytesPath.Text = pPathTableData;


            //
            _mtCreater = new TableDataLoaderCreater();
            _sb = new StringBuilder();

            RefreshFileInfosWithList();
        }

        /// <summary>
        /// 
        /// </summary>
        void RefreshFileInfosWithList()
        {
            //
            RefreshFileInfos();

            //
            listBox1.SelectionMode = SelectionMode.None;
            listBox1.Items.Clear();
            foreach (var item in _fileInfos)
            {
                listBox1.Items.Add(item.Name);
            }
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
            //
            RefreshFileInfosWithList();

            //
            label1.Text = $"Refresh complete";
        }

        private void OnBtn_BrowseBytesPath(object sender, EventArgs e)
        {
            var selectedPath = SelectOutputFolder(pPathTableData, ".bytes 저장 경로 선택");
            if (selectedPath == null)
                return;

            pPathTableData = selectedPath;
            pPathTableLocalData = selectedPath;
            textBoxBytesPath.Text = selectedPath;
        }

        private void OnBtn_BrowseScriptPath(object sender, EventArgs e)
        {
            var selectedPath = SelectOutputFolder(pPathScript, ".cs 저장 경로 선택");
            if (selectedPath == null)
                return;

            pPathScript = selectedPath;
            pPathGlobalData = selectedPath;
            textBoxScriptPath.Text = selectedPath;
        }

        static string? SelectOutputFolder(string initialPath, string description)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = description,
                SelectedPath = Directory.Exists(initialPath)
                    ? initialPath
                    : Directory.GetCurrentDirectory(),
                ShowNewFolderButton = true,
                UseDescriptionForTitle = true
            };

            return dialog.ShowDialog() == DialogResult.OK
                ? dialog.SelectedPath
                : null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtn_Confirm(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = false;
            label1.Text = $"Converting...";

            try
            {
                Directory.CreateDirectory(pPathScript);
                Directory.CreateDirectory(pPathTableData);

                var classNames = new List<string>();

                foreach (var item in _fileInfos)
                {
                    var fileName = Path.GetFileNameWithoutExtension(item.Name);

                    using var snapshot = CreateWorkbookSnapshot(item.FullName);
                    using var workBook = new XLWorkbook(snapshot);

                    var numberName = fileName.Substring(1, 3);
                    if (numberName[0] == '0')
                    {
                        CreateEnum(fileName, workBook);
                    }
                    else
                    {
                        CreateClass(fileName, workBook);
                        CreateData(fileName, workBook, numberName == "999");
                    }

                    foreach (var worksheet in workBook.Worksheets)
                    {
                        classNames.Add(worksheet.Name == "Data"
                            ? fileName
                            : $"{fileName}_{worksheet.Name}");
                    }
                }

                _mtCreater.Create(classNames);
                label1.Text = $"Convert complete";
            }
            catch (Exception exception)
            {
                label1.Text = $"Convert failed";
                MessageBox.Show(
                    exception.Message,
                    "Convert failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                button1.Enabled = true;
                button2.Enabled = true;
            }
        }

        /// <summary>
        /// Excel에서 열어 둔 파일도 마지막으로 저장된 내용을 읽을 수 있도록
        /// 공유 읽기 모드로 메모리 스냅샷을 만든다.
        /// </summary>
        static MemoryStream CreateWorkbookSnapshot(string path)
        {
            const int maxAttempts = 3;

            for (var attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    using var source = new FileStream(
                        path,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite | FileShare.Delete);

                    var snapshot = new MemoryStream();
                    source.CopyTo(snapshot);
                    snapshot.Position = 0;
                    return snapshot;
                }
                catch (IOException) when (attempt < maxAttempts)
                {
                    Thread.Sleep(100);
                }
            }

            throw new IOException($"Excel 파일을 읽을 수 없습니다: {path}");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="className"></param>
        /// <param name="workBook"></param>
        void CreateEnum(string className, XLWorkbook workBook)
        {
#if !DEBUG
            //
            var fs = new FileStream($"{pPathGlobalData}\\{className}.cs", FileMode.Create, FileAccess.Write);
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
            var tempVariables = new List<string>();



            for (int col = 1; col <= columnCount; col++)
            {
                for (int row = 2; row <= rowCount; row++)
                {
                    var cellValue = worksheet.Cell(row, col).Value;
                    if (cellValue.IsText)
                    {
                        //
                        var text = worksheet.Cell(row, col).GetText();
                        if (string.IsNullOrEmpty(text))
                            continue;

                        //
                        _sb.AppendFormat(row == 2 ? "public enum E{0}\r\n{{\r\n    None,\r\n" : "    {0},\r\n", text);
                    }
                }
                _sb.Append("    End\r\n");
                _sb.Append("}\r\n");

                if (col != columnCount)
                {
                    _sb.Append("\r\n");
                }
            }

#if !DEBUG
            sw.Write(_sb);
            sw.Close();
            fs.Close();
#endif
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="className"></param>
        /// <param name="workBook"></param>
        void CreateClass(string fileName, XLWorkbook workBook)
        {
            //
            var totalSheetCount = workBook.Worksheets.Count;
            IXLWorksheet? worksheet = null;
            IXLRange? range = null;
            for (int i = 0; i < totalSheetCount; i++)
            {
                //
                worksheet = workBook.Worksheet(i + 1);
                if (worksheet == null)
                    continue;

                range = worksheet.RangeUsed();

                //
                var className = worksheet.Name == "Data" ? fileName : $"{fileName}_{worksheet.Name}";
#if !DEBUG                
                var fs = new FileStream($"{pPathScript}\\{className}.cs", FileMode.Create, FileAccess.Write);
                var sw = new StreamWriter(fs);
#endif

                //
                if (range == null)
                    return;

                var tempVariables = new List<KeyValuePair<string, string>>();
                var arrayKeys = new List<string>();
                var isArray = false;
                for (int col = 1;  col <= range.ColumnCount();    col++)
                {
                    //                    
                    isArray = false;
                    var variableName = worksheet.Cell(2, col).GetText();

                    //
                    if (variableName.Substring(0, 1) == ".")
                        continue;

                    //
                    if (variableName.Substring(0, 1) == "[")
                    {
                        if (arrayKeys.Contains(variableName))
                        {
                            continue;
                        }
                        else
                        {
                            arrayKeys.Add(variableName);
                            variableName = variableName.Substring(1, variableName.Length - 2);
                            isArray = true;
                        }
                    }

                    //
                    var dataType = worksheet.Cell(3, col).Value.GetText();
                    dataType += isArray ? "[]" : "";

                    //
                    var temp = new KeyValuePair<string, string>(dataType, variableName);
                    tempVariables.Add(temp);
                }
                var data = ClassCode(className, tempVariables);
#if !DEBUG
                sw.Write(data);
                sw.Close();
                fs.Close();
#endif
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="variables"></param>
        /// <returns></returns>
        string ClassCode(string className, List<KeyValuePair<string, string>> variables)
        {
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
        /// 엑셀 워크북의 각 시트를 JSON 파일로 변환한다.
        /// 2행: 변수명
        /// 3행: 자료형
        /// 4행 이후: 데이터
        /// [name] 형태로 반복되는 열은 JSON 배열로 변환한다.
        /// </summary>
        void CreateData(string fileName, XLWorkbook workBook, bool isLocal)
        {
            var path = isLocal ? pPathTableLocalData : pPathTableData;

            foreach (var worksheet in workBook.Worksheets)
            {
                var range = worksheet.RangeUsed();

                if (range == null)
                    continue;

                var rowCount = range.LastRow().RowNumber();
                var columnCount = range.LastColumn().ColumnNumber();

                var className = worksheet.Name == "Data"
                    ? fileName
                    : $"{fileName}_{worksheet.Name}";

                var root = new JsonArray();

                for (int row = 4; row <= rowCount; row++)
                {
                    // 완전히 비어 있는 행은 제외
                    if (worksheet.Row(row).Cells(1, columnCount)
                        .All(cell => cell.IsEmpty()))
                    {
                        continue;
                    }

                    var jsonObject = new JsonObject();

                    for (int col = 1; col <= columnCount;)
                    {
                        var header = worksheet.Cell(2, col).GetText().Trim();

                        // 헤더가 없거나 "."으로 시작하는 열은 제외
                        if (string.IsNullOrEmpty(header) ||
                            header.StartsWith(".", StringComparison.Ordinal))
                        {
                            col++;
                            continue;
                        }

                        if (TryGetArrayName(header, out var arrayName))
                        {
                            var array = new JsonArray();
                            var arrayTypeName = string.Empty;
                            var typeName = worksheet.Cell(3, col).GetText().Trim();                            

                            while (col <= columnCount)
                            {
                                var currentHeader = worksheet.Cell(2, col).GetText().Trim();

                                if (!TryGetArrayName(currentHeader, out var currentArrayName) ||
                                    !string.Equals(arrayName, currentArrayName, StringComparison.Ordinal))
                                {
                                    break;
                                }

                                var declaredTypeName =
                                    worksheet.Cell(3, col).GetText().Trim();

                                // 자료형이 선언된 열에서만 갱신한다.
                                // 이후 빈칸은 이전 자료형을 그대로 사용한다.
                                if (!string.IsNullOrWhiteSpace(declaredTypeName))
                                {
                                    arrayTypeName = declaredTypeName;
                                }

                                if (string.IsNullOrWhiteSpace(arrayTypeName))
                                {
                                    throw new InvalidDataException(
                                        $"{worksheet.Name} 시트의 {col}열 배열 자료형을 알 수 없습니다.");
                                }

                                array.Add(ConvertCellToJson(
                                    worksheet.Cell(row, col),
                                    arrayTypeName));

                                col++;
                            }

                            jsonObject[arrayName] = array;
                            continue;
                        }

                        var scalarTypeName =
                            worksheet.Cell(3, col).GetText().Trim();

                        jsonObject[header] = ConvertCellToJson(
                            worksheet.Cell(row, col),
                            scalarTypeName);

                        col++;
                    }

                    root.Add(jsonObject);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                var json = root.ToJsonString(options);

#if !DEBUG
        Directory.CreateDirectory(path);

        var outputPath = Path.Combine(path, $"{className}.bytes");

        File.WriteAllText(
            outputPath,
            json,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
#endif
            }
        }

        /// <summary>
        /// "[rewardKeys]" 형태의 헤더인지 확인하고
        /// 실제 JSON 속성명인 "rewardKeys"를 반환한다.
        /// </summary>
        static bool TryGetArrayName(string header, out string arrayName)
        {
            arrayName = string.Empty;

            if (header.Length < 3 ||
                header[0] != '[' ||
                header[^1] != ']')
            {
                return false;
            }

            arrayName = header[1..^1].Trim();

            return arrayName.Length > 0;
        }

        /// <summary>
        /// 엑셀 셀을 3행에 정의된 자료형에 맞춰 JSON 값으로 변환한다.
        /// </summary>
        static JsonNode? ConvertCellToJson(
    IXLCell cell,
    string typeName)
        {
            if (cell.IsEmpty())
                return null;

            if (string.IsNullOrWhiteSpace(typeName))
            {
                throw new InvalidDataException(
                    $"자료형이 없습니다. 셀 위치: {cell.Address}");
            }

            var normalizedType = typeName.Trim().ToLowerInvariant();

            switch (normalizedType)
            {
                case "byte":
                    if (cell.TryGetValue<byte>(out var byteValue))
                        return JsonValue.Create(byteValue);
                    break;

                case "short":
                case "int16":
                    if (cell.TryGetValue<short>(out var shortValue))
                        return JsonValue.Create(shortValue);
                    break;

                case "int":
                case "int32":
                    if (cell.TryGetValue<int>(out var intValue))
                        return JsonValue.Create(intValue);
                    break;

                case "long":
                case "int64":
                    if (cell.TryGetValue<long>(out var longValue))
                        return JsonValue.Create(longValue);
                    break;

                case "float":
                case "single":
                    if (cell.TryGetValue<float>(out var floatValue))
                        return JsonValue.Create(floatValue);
                    break;

                case "double":
                    if (cell.TryGetValue<double>(out var doubleValue))
                        return JsonValue.Create(doubleValue);
                    break;

                case "decimal":
                    if (cell.TryGetValue<decimal>(out var decimalValue))
                        return JsonValue.Create(decimalValue);
                    break;

                case "bool":
                case "boolean":
                    if (cell.TryGetValue<bool>(out var boolValue))
                        return JsonValue.Create(boolValue);

                    if (cell.TryGetValue<int>(out var boolNumber))
                    {
                        if (boolNumber == 1)
                            return JsonValue.Create(true);

                        if (boolNumber == 0)
                            return JsonValue.Create(false);
                    }
                    break;

                case "string":
                    return JsonValue.Create(
                        cell.GetFormattedString(
                            CultureInfo.InvariantCulture));
            }

            /*
             * EMissionType처럼 int, long 등이 아닌 자료형은
             * enum 또는 문자열로 간주한다.
             */
            var stringValue = cell.GetFormattedString(
                CultureInfo.InvariantCulture);

            if (string.IsNullOrWhiteSpace(stringValue))
                return null;

            return JsonValue.Create(stringValue.Trim());
        }
    }
}
