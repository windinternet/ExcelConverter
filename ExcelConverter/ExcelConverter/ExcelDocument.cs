using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
namespace ExcelConverter
{
    public class ExcelDocument<T> :IEnumerable<T> where T :class
    {
        StringBuilder TableBuilder { get; set; }
        Dictionary<int, int> PositionMapping { get; set; }
        int Index { get; set; }
        List<T> DataList { get; set; }
        public Exception InnerException { get; private set; }
        PropertyInfo[] Properties { get; set; }
        public bool HasRows { get; private set; }
        public ExcelDocument(bool containerHeader=true)
        {
            InitProperty();
            if (containerHeader)
            {
                InitHeader();
            }
        }
        private void InitHeader()
        {
            TableBuilder.AppendLine(GetHeader());
        }
        public string GetHeader()
        {
            StringBuilder headerBuilder = new StringBuilder();
            headerBuilder.AppendLine("<tr>");
            var type = typeof(NameAttribute);
            foreach (var property in Properties)
            {
                if (Reflection.ContainsAttribute(property, type))
                {
                    var name = (NameAttribute)Reflection.GetAttribute(property, type);
                    headerBuilder.AppendLine(string.Format("<th >{0}</th>", name.Name));
                }
                else
                {
                    headerBuilder.AppendLine(string.Format("<th >{0}</th>", property.Name));
                }
            }
            headerBuilder.AppendLine("</tr>");
            return headerBuilder.ToString();
        }
        public string ReadRowByString()
        {
            var model = ReadRow();
            return GetRowString(model);
        }
        public ExcelDocument(string tableName)
        {
            InitProperty();
            TableBuilder.AppendLine("<tr>");
            if (string.IsNullOrWhiteSpace(tableName))
            {
                tableName = typeof(T).Name;
            }
            TableBuilder.AppendLine(string.Format("<td  align='center' colspan='{0}'>{1}</td>", Properties.Length.ToString(), tableName));
            TableBuilder.AppendLine("</tr>");
            InitHeader();
        }
        private void InitProperty()
        {
            TableBuilder = new StringBuilder();
            DataList = new List<T>();
            Properties = Reflection.GetProperties<T>();
            TableBuilder.AppendLine("<table border='1'>");
            PositionMapping = new Dictionary<int, int>();
        }
        public void WriteRow(T model,bool error=true)
        {
            if (model== null && error)
            {
                throw new ArgumentNullException("添加空行通常是无效的，如果当真有业务需求，请将error参数设置为false，默认为true");
            }
            var position = TableBuilder.Length - 1;
            PositionMapping.Add(DataList.Count, position);
            DataList.Add(model);
            HasRows = true;
            TableBuilder.Append(GetRowString(model));
        }
        public void ResetIndex()
        {
            Index = 0;
        }
        private string GetRowString(T model)
        {
            StringBuilder rowBuilder = new StringBuilder();
            rowBuilder.AppendLine("<tr>");
            if (model == null)
            {
                for (int i = 0; i < Properties.Length; i++)
                {
                    rowBuilder.AppendLine("<td ></td>");
                }
            }
            else
            {
                for (int i = 0; i < Properties.Length; i++)
                {
                    rowBuilder.AppendLine(string.Format("<td >{0}</td>",Properties[i].GetValue(model,null)));
                }
            }
            rowBuilder.AppendLine("</tr>");
            return rowBuilder.ToString();
        }
        private string End()
        {
            return string.Concat("</table>",Environment.NewLine);
        }
        public void RemoveRow(int index)
        {
            if (index < 0 || index >= DataList.Count)
                throw new IndexOutOfRangeException(string.Format("所提供的的索引超出了行数的范围，当前共{0}行", DataList.Count.ToString()));
            DataList.RemoveAt(index);
            int len = 0;
            if (PositionMapping.ContainsKey(index + 1))
            {
                len = PositionMapping[index + 1] - PositionMapping[index];
                TableBuilder.Remove(PositionMapping[index], len);
                var keys=PositionMapping.Keys.ToList();
                foreach (int i in keys)
                {
                    if (PositionMapping.ContainsKey(i + 1))
                    {
                        PositionMapping[i] = PositionMapping[i + 1] - len;
                    }
                }
                PositionMapping.Remove(DataList.Count);
            }
            else
            {
                len = TableBuilder.Length - PositionMapping[index] - 1;
                TableBuilder.Remove(PositionMapping[index], len);
                PositionMapping.Remove(index);
            }
            if (Index >= DataList.Count)
                Index = DataList.Count - 1;
        }
        public void SetPosition(int i)
        {
            if (i < 0 || i >= DataList.Count)
                throw new IndexOutOfRangeException(string.Format("所提供的的索引超出了行数的范围，当前共{0}行",DataList.Count.ToString()));
            Index = i;
        }
        private void Move()
        {
            if (Index >= DataList.Count-1)
            {
                HasRows = false;
            }
            else
            {
                HasRows = true;
                Index++;
            }
        }
        public T ReadRow()
        {
            T model= Index>=DataList.Count?null:DataList[Index];
            Move();
            return model;
        }
        public IEnumerator<T> GetEnumerator()
        {
            return DataList.GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return DataList.GetEnumerator();
        }
        public string ToHtml()
        {
            return TableBuilder.ToString() + End();
        }
        public static ExcelDocument<T> LoadFromString<T>(string html) where T:class
        {
            ExcelDocument<T> excel = new ExcelDocument<T>();
            try
            {
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                string head = null;
                if (html == null)
                {
                    excel.InnerException = new ArgumentNullException("参数html为空，因此未能解析到数据！");
                    return excel;
                }
                head = html.Substring(0, html.IndexOf(">") + 1);
                Regex reg = new Regex(@"<table(\s|\S)*>(\s|\S)*</table>");
                html = reg.Match(html).Value;
                Regex reg2 = new Regex(@"<!\[(\s|\S)*\]>(\s|\S)*<!\[(\s|\S)*\]>");
                html = reg2.Replace(html, "");
                Regex attrReg = new Regex(@"\S*=[a-zA-Z0-9]+");
                html = attrReg.Replace(html, new MatchEvaluator((Match match) =>
                {
                    try
                    {
                        if (match.Value.Contains("'") || match.Value.Contains("\""))
                        {
                            return match.Value;
                        }
                        var strLeft = match.Value.Substring(0, match.Value.IndexOf("="));
                        var strRight = match.Value.Substring(match.Value.IndexOf("=") + 1);
                        return string.Format("{0}=\"{1}\"", strLeft, strRight);
                    }
                    catch (Exception)
                    {
                        return "";
                    }
                }));
                var table = html.Substring(0, html.IndexOf(">") + 1);
                var tr = html.Substring(html.IndexOf("<tr"));
                html = table + tr;
                html = head + html + "</html>";
                doc.LoadXml(html);
                var rows = doc.ChildNodes[0].ChildNodes[0].ChildNodes;
                int i = 0;
                var propertise = Reflection.GetProperties<T>();
                var sumCol = 0;
                foreach (System.Xml.XmlNode row in rows)
                {
                    sumCol += row.ChildNodes.Count;
                }
                foreach (System.Xml.XmlNode row in rows)
                {
                    if (!row.Name.Equals("tr") || row.ChildNodes.Count<sumCol/rows.Count)
                        continue;
                    try
                    {
                        var model = Reflection.CreateInstance<T>(null);
                        if (model == null)
                        {
                            excel.InnerException = new ArgumentException(string.Format("T:{0}不是一个可以实例化的类型，请适当的修改原代码。",typeof(T).FullName));
                            return excel;
                        }
                        i = 0;
                        foreach (System.Xml.XmlNode td in row .ChildNodes)
                        {
                            Reflection.SetPropertyValue(model,propertise[i], td.InnerText);
                            i++;
                        }
                        excel.WriteRow(model, false);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
                return excel;
            }
            catch (Exception e)
            {
                excel.InnerException = e;
                return excel;
            }
        }
        public static ExcelDocument<T> LoadFromFile<T>(string filepath) where T : class
        {
            if (!File.Exists(filepath))
                throw new FileNotFoundException(string.Format("未能找到：{0}文件",filepath));
            var reader = new StreamReader(filepath);
            var html =reader.ReadToEnd();
            reader.Close();
            reader.Dispose();
            return LoadFromString<T>(html);
        }
    }
}
