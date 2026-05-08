using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableDataConverter
{
    internal class DefineCreater
    {
        StringBuilder _sb;

        public void Create(string className, XLWorkbook workBook)
        {
            _sb ??= new StringBuilder();
            _sb.Clear();
            _sb.Append($"using System;\r\nSystem.Collections;\r\nusing System.Collections.Generic;\r\nUnityEngine;\r\n");
            
        }
    }
}
