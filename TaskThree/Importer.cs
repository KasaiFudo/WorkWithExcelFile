using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TaskThree
{
    class Importer
    {
        public List<KeyValuePair<string, string>> mappings;
        public string FilePath { get; set; }

        public Importer(string filePath)
        {
            FilePath = filePath;
            mappings = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("Код товара","ProductID"),
                new KeyValuePair<string, string>("Наименование","Name"),
                new KeyValuePair<string, string>("Ед. измерения","UnitOfMeasure"),
                new KeyValuePair<string, string>("Цена товара за единицу","Price"),
                new KeyValuePair<string, string>("Код клиента","CustomerID"),
                new KeyValuePair<string, string>("Наименование организации","OrganizationName"),
                new KeyValuePair<string, string>("Адрес","Adress"),
                new KeyValuePair<string, string>("Контактное лицо (ФИО)","Contact"),
                new KeyValuePair<string, string>("Код заявки","RequestID"),
                new KeyValuePair<string, string>("Номер заявки","Number"),
                new KeyValuePair<string, string>("Требуемое количество","ProductAmount"),
                new KeyValuePair<string, string>("Дата размещения","Date"),
            };
        }

        public List<T> ImportExel<T>(string sheetName)
        {
            try
            {
                List<T> list = new();
                Type typeOfObjects = typeof(T);
                using (IXLWorkbook workbook = new XLWorkbook(FilePath))
                {
                    var worksheet = workbook.Worksheets.Where(w => w.Name == sheetName).First();
                    var properties = typeOfObjects.GetProperties();
                    var headerInfo = worksheet.FirstRow().Cells().Select((v, i) => new { ColName = v.Value, ColIndex = i + 1 });

                    foreach (IXLRow row in worksheet.RowsUsed().Skip(1))
                    {
                        T obj = (T)Activator.CreateInstance(typeOfObjects);

                        foreach (var prop in properties)
                        {
                            var mapping = mappings.SingleOrDefault(m => m.Value == prop.Name);
                            var colName = mapping.Key;
                            var colIndex = headerInfo.SingleOrDefault(s => s.ColName.ToString() == colName).ColIndex;
                            var val = row.Cell(colIndex).Value;
                            var type = prop.PropertyType;
                            object? value = null;

                            if (val.Type == XLDataType.Text)
                            {
                                value = val.ToString();
                            }
                            else if (val.Type == XLDataType.Number)
                            {
                                value = Convert.ChangeType(val.GetNumber(), type);
                            }
                            else if (val.Type == XLDataType.DateTime)
                            {
                                value = Convert.ChangeType(val.GetDateTime(), type);
                            }
                            prop.SetValue(obj, value);
                        }
                        list.Add(obj);
                    }
                }

                return list;
            } catch (NullReferenceException)
            {
                Console.WriteLine("Не найдена таблица с именем: " + sheetName);
                return null;
            }

        }
        public void ExportExel<T>(string nameSheet, List<T> newSheet)
        {
            using (IXLWorkbook workbook = new XLWorkbook(FilePath))
            {
                //var rusHeaderInfo = workbook.Worksheets.Where(w => w.Name == "Клиенты")
                //    .First().FirstRow().Cells().Select((v, i) => new { ColName = v.Value, ColIndex = i + 1 });
                workbook.Worksheets.Where(w => w.Name == nameSheet).First().Delete();
                var worksheet = workbook.AddWorksheet(nameSheet).FirstCell().InsertTable(newSheet, false);
                foreach( var cell in worksheet.FirstRow().Cells() )
                {
                    var map = mappings.FirstOrDefault(m => m.Value == cell.Value.ToString());
                    cell.Value = map.Key;
                }
                workbook.SaveAs(FilePath);
            }
        }
    }
}
