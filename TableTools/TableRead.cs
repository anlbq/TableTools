using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.IO;
using Newtonsoft.Json;
namespace TableTools
{
   public class TableRead
    {
       public  static string JsonSavepath = "/table";
        static string fliename0=".xlsx";
   
        public static void Init() {

            string path = Directory.GetCurrentDirectory();
            JsonSavepath = path + JsonSavepath;

            string[] files = Directory.GetFiles(path);
            for (int i=0; i<files.Length;i++) {
                if (files[i].IndexOf(fliename0) < 0)continue;

                string file = files[i];
                ExcelPackage pack = new ExcelPackage();
                FileStream streamReader = File.OpenRead(file);
                pack.Load(streamReader);
                int index = streamReader.Name.IndexOf(fliename0);
                string spath = streamReader.Name.Remove(index);
                if (Directory.Exists(spath))
                {
                    Directory.Delete(spath,true);
                }
                Directory.CreateDirectory(spath);
               
                List<SheetInfo> sheetInfos = new List<SheetInfo>();              
                foreach (ExcelWorksheet worksheet in pack.Workbook.Worksheets) {
                    int endrow = worksheet.Cells.End.Row;
                    int endcol = worksheet.Cells.End.Column;
                    SheetInfo sheet = new SheetInfo();
                    sheet.SheetName = worksheet.Name;
                    sheet.Savepath = spath;
                    for (int j = worksheet.Cells.Start.Row; j < endrow; j++) {
                        cellInfo cinfo = new cellInfo();
                        if (worksheet.Cells[j, 1].Value == null) break;
                        for (int k = worksheet.Cells.Start.Column; k < endcol; k++) {
                            if (worksheet.Cells[j, k].Value == null)
                            {
                                break;
                            }
                            else {
                                cinfo.info.Add(worksheet.Cells[j, k].Value.ToString());
                            }
                        }
                        sheet.CellInfos.Add(cinfo);
                    }
                    sheetInfos.Add(sheet);
                }
                          
                for (int j = 0; j < sheetInfos.Count; j++) {                
                    sheetInfos[j].SerializeObject();
                }
            }
        }
    }



    public class cellInfo {
       public List<string> info = new List<string>(20);
    }

    public class SheetInfo
    {
        public string SheetName;
        public List<cellInfo> CellInfos = new List<cellInfo>();
        public string Jsoninfo;
        public string ClassInfo;

        public string Savepath;
        static List<string> head=new List<string>() { "using System;", "using System.Collections.Generic;" };
        static string baseClass=" public class {0}";
        static string  baseClassinfo = "{#}";

        static string ClassFileName = ".cs";
        static string JsonFileName = ".json";

        public void SerializeObject()
        {
            if (CellInfos.Count>3) {
                Jsoninfo = GetJsonInfo();
                Console.Write(Jsoninfo);
                ClassInfo = GetClassInfo();
                Console.Write(ClassInfo);
                SaveInfo();
            }
        }

        private void SaveInfo()
        {
            string spath = Savepath;      
            string jsonpath = spath+"/"+SheetName+ JsonFileName;
            StreamWriter streamWriter;
            if (File.Exists(jsonpath)) {
                streamWriter = File.AppendText(jsonpath);
            }
            else {
                streamWriter = File.CreateText(jsonpath);
            }          
            byte[] data=Encoding.UTF8.GetBytes(Jsoninfo);
            streamWriter.Write(Jsoninfo);
            streamWriter.Close();
            streamWriter.Dispose();
            string classpath= spath + "/" + SheetName + ClassFileName;
            if (File.Exists(classpath))
            {
                streamWriter = File.AppendText(classpath);
            }
            else
            {
                streamWriter = File.CreateText(classpath);
            }
            data = Encoding.UTF8.GetBytes(ClassInfo);
            streamWriter.Write(ClassInfo);
            streamWriter.Close();
            streamWriter.Dispose();
        }


        private string GetJsonInfo() {
       
            string text = string.Empty;        
            cellInfo numtype = CellInfos[1];
            cellInfo numname = CellInfos[2];
            for (int i=3;i<CellInfos.Count;i++)
            {
                for (int j=0;j<CellInfos[i].info.Count;j++) {
                    text += CellInfos[i].info[j].SerializeJson(numname.info[j],numtype.info[j]);
                    if (j< CellInfos[i].info.Count-1)
                    {
                        text += ",";
                    }
                }
            }
            text = "{"+text+"}";
            return text;
        }

        private string GetClassInfo()
        {
            string cinfo = string.Empty;
            for (int i = 0; i < head.Count; i++)
            {
                cinfo += head[i] + "\r\n";
            }
            string text = string.Empty;
            cellInfo numinfo = CellInfos[0];
            cellInfo numtype = CellInfos[1];
            cellInfo numname = CellInfos[2];
            for (int i=0;i<numinfo.info.Count;i++) {
                string info = numinfo.info[i].SerializeClass(numname.info[i], numtype.info[i]);
                text += info + "\r\n";
            }
            text =string.Format(baseClass,SheetName)+ baseClassinfo.Replace("#",text);
            cinfo += text;
            return cinfo;
        }

    }

    public static class SerializeObject_Exend {
        static string basestr = "public {0} {1}";//"public {0} {1} {get;set;}// {2}";
        static string id = "\"{0}\":";
        static string strlist = "\"{0}\"";

        public static string SerializeJson(this string value,string name,string vtype){
          
            string text = "";
            name = string.Format(id,name);
            text += name;
            vtype = vtype.ToLower();
            string v = "";
            if (vtype.Equals(Tag.STRING))
            {
                v = "\"" + value + "\"";
            }
            else if (vtype.Equals(Tag.STRINGLIST))
            {
                string lists = "[{0}]";
                string v1 = "";
                string[] strs = value.Split(",");
                for (int i = 0; i < strs.Length; i++) {
                    v1 += string.Format(strlist, strs[i]);
                    if (i < strs.Length - 1) {
                        v1 += ",";
                    }
                }
                v = string.Format(lists, v1);
            } else if (vtype.IndexOf("[,]")>0) {
                if (vtype.IndexOf(Tag.BOOL) >= 0) value = value.ToLower();
                string lists = "[{0}]";
                string v1 = "";
                string[] strs = value.Split(",");
                for (int i = 0; i < strs.Length; i++)
                {
                    v1 += strs[i];
                    if (i < strs.Length - 1)
                    {
                        v1 += ",";
                    }
                }
                v = string.Format(lists, v1);
            }
            else {
                if (vtype.IndexOf(Tag.BOOL) >= 0) value = value.ToLower();
                v = value;
            }
            text += v;
            return text;
        }


        public static string SerializeClass(this string info, string name, string vtype) {
            if (vtype.Equals(Tag.INTLIST))
            {
                vtype = Tag.Listint;
            }
            else if (vtype.Equals(Tag.FLOATLIST)) {
                vtype = Tag.Listfloat;
            }
            else if (vtype.Equals(Tag.DOUBLELIST))
            {
                vtype = Tag.Listdouble;
            }
            else if (vtype.Equals(Tag.BOOLLIST))
            {
                vtype = Tag.Listbool;
            }
            else if (vtype.Equals(Tag.STRINGLIST))
            {
                vtype = Tag.Liststring;
            }
            return string.Format(basestr,vtype,name)+ " {get;set;}//"+info;
        }




    }

    public class Tag {
        public const string INT = "int";
        public const string FLOAT = "float";
        public const string DOUBLE = "double";
        public const string BOOL = "bool";
        public const string STRING = "string";
        public const string INTLIST = "int[,]";
        public const string FLOATLIST = "float[,]";
        public const string DOUBLELIST = "double[,]";
        public const string BOOLLIST = "bool[,]";
        public const string STRINGLIST = "string[,]";


        public const string Listint = "List<int>";
        public const string Listfloat = "List<float>";
        public const string Listdouble = "List<double>";
        public const string Listbool = "List<bool>";
        public const string Liststring = "List<string>";
    }

    

}
