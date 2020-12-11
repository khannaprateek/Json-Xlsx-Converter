using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using Syncfusion.XlsIO;
using Newtonsoft.Json;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace JsontoExcel
{
    class KeyValueModel
    {

        public string Key { get; set; }
        public dynamic Value { get; set; }

        public KeyValueModel() { }
        public KeyValueModel(string Key, dynamic Value)
        {
            this.Key = Key;
            this.Value = Value;
        }

    }
    class Program
    {
        public static JObject ReadFromJsonFile(string filePath)
        {
            using (StreamReader file = System.IO.File.OpenText(filePath))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                JObject jsonObj = (JObject)JToken.ReadFrom(reader);

                return jsonObj;
            }
        }
       
        public static void ConvertJsonToXlsx(string filePath,string fileName)
        {
            string outputFile = filePath.Replace(".json", ".xlsx");
            //string filePath = Directory.GetCurrentDirectory();
            //filePath += fileName;

            //Console.WriteLine(filePath);

            JObject jsonObj = ReadFromJsonFile(filePath);
            Dictionary<string, dynamic> dict = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(jsonObj.ToString());

            Dictionary<string, string> output = new Dictionary<string, string>();
            output = FlattenTheDictionary(dict, output);

            WriteToXlsx(output,outputFile);
        }

        public static void ConvertXlsxToJson(string filePath,string fileName)
        {

            string outputFile = filePath.Replace("xlsx", "json");
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Open(filePath, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];
                Dictionary<string, List<KeyValueModel>> unflattenedDictionary = new Dictionary<string, List<KeyValueModel>>();
                Dictionary<string, KeyValueModel> collectionObject = new Dictionary<string, KeyValueModel>();

                int rowCount = worksheet.UsedRange.LastRow;
                int colCount = worksheet.UsedRange.LastColumn;
                string key;
                KeyValueModel value;
                List<string> keyList;

                for (int row = 1; row <= rowCount; row++)
                {
                    keyList = worksheet.Range[row, 1].Text.Split('.').ToList();
                    UnflattenTheJson(ref collectionObject, keyList, worksheet.Range[row, 2].Text, keyList.Count);

                    key = collectionObject.Keys.ElementAt(0);
                    value = collectionObject[key];

                    if (unflattenedDictionary.ContainsKey(key))
                    {
                        unflattenedDictionary[key].Add(createKeyValueModelObject(value.Key, value.Value));
                    }
                    else
                    {
                        unflattenedDictionary.Add(key, new List<KeyValueModel>() { value });
                    }

                }
                Dictionary<string, dynamic> finalDict = new Dictionary<string, dynamic>();
                foreach (var item in unflattenedDictionary)
                {

                    finalDict.Add(item.Key, item.Value.ToDictionary(eachItem => eachItem.Key, eachItem => eachItem.Value));
                }

                string jsonString = JsonConvert.SerializeObject(finalDict);

               
                File.WriteAllText(outputFile, jsonString);
                Console.WriteLine($"Your file has been successfully converted to \n{outputFile} ");
            }
        }


        public static Dictionary<string, string> FlattenTheDictionary(
            Dictionary<string, dynamic> dict,
            Dictionary<string, string> output,
            string parentKey = "")
        {
            if (dict.Count() == 0)
            {
                return output;
            }

            foreach (var item in dict)
            {
                if (item.Value.GetType() == typeof(string))
                {
                    output.Add(parentKey + "." + item.Key, item.Value);
                }
                else if (item.Value.GetType() == typeof(JObject))
                {
                    FlattenTheDictionary(JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(item.Value.ToString()),
                        output, parentKey != "" ? parentKey + "." + item.Key : item.Key);
                }

            }

            return output;

        }
        public static void WriteToXlsx(Dictionary<string, string> output,string fileName)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                worksheet.SetColumnWidth(1, 75);
                worksheet.SetColumnWidth(2, 105);
                worksheet.SetRowHeight(1, 20);
                int row = 1;
                foreach (var item in output)
                {
                    worksheet.Range[row, 1].Text = item.Key;
                    worksheet.Range[row, 2].Text = item.Value;
                    row++;
                }
                Console.WriteLine($"File has been successfully converted...Added {row - 1} rows to {fileName} file");
                //Saving the workbook
                workbook.SaveAs(fileName);
                workbook.Close();
                excelEngine.Dispose();
            }

        }
        public static KeyValueModel UnflattenTheJson(ref Dictionary<string, KeyValueModel> dict, List<string> keyList, string value, int count = 0)
        {
            var copyOfKeylist = keyList.ToList();
            copyOfKeylist.RemoveAt(0);
            if (keyList.Count == 1)
            {
                return createKeyValueModelObject(keyList[0], value);
            }
            else if (keyList.Count == count)
            {
                dict = new Dictionary<string, KeyValueModel>() { { keyList[0], UnflattenTheJson(ref dict, copyOfKeylist, value) } };
            }

            return new KeyValueModel(keyList[0], UnflattenTheJson(ref dict, copyOfKeylist, value));

        }
        public static KeyValueModel createKeyValueModelObject(string key, dynamic value)
        {
            return new KeyValueModel(key, value);
        }

        static void Main(string[] args)
        {
            if( args.Length == 0)
            {
                Console.WriteLine("Please provide either a json or a xlsx file!!");
            }
            else
            {
                var startDirectory = Directory.GetCurrentDirectory();
                
                string filePath = startDirectory + "\\" + args[0];
                FileInfo fileinfo = new FileInfo(filePath);
                Console.WriteLine(fileinfo);
                string fileName = fileinfo.Name.Replace(fileinfo.Extension,"");
                
                if (fileinfo.Extension == ".xlsx")
                {
                    Console.WriteLine($"Converting your {fileName+".xlsx"} to json file...");
                    ConvertXlsxToJson(filePath,fileName);
                }
                else if (fileinfo.Extension == ".json")
                {
                    Console.WriteLine($"Converting your {fileName+".json"} to xlsx file...");
                    ConvertJsonToXlsx(filePath,fileName);
                }
                else 
                {
                    Console.WriteLine("Please provide either a json or a xlsx file!!");

                }

            }
        }
    }
}

