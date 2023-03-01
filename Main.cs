using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPIReadingDataFromExcelFile
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            //вызывается диалоговое окно
            OpenFileDialog openFileDialog1 = new OpenFileDialog // создаем переменную
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop), // задается директория по умолчанию, которая будет открываться при выводе окна
                Filter = "Excel files (*.xlsx) | *.xlsx" // указывается фильтр, указали фильтр только для файла Excel
            };

            string filePath = string.Empty;

            if(openFileDialog1.ShowDialog() == DialogResult.OK) // если все прошло удачно
            {
                filePath = openFileDialog1.FileName; //то сохраняем путь к файлу в переменную filePath
            }

            if (string.IsNullOrEmpty(filePath)) //если путь не указан 
                return Result.Cancelled; //возвращаем отмену

            //собираем все помещения с которыми будем работать
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            //мы должны открыть по указанному пути файл Excel и прочитать из него данные, т.е открываем файл не создаем новый и открываем этот файл для чтения
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(filePath);//создаем переменную для чтения файла Excel, указываем ему путь
                ISheet sheet = workbook.GetSheetAt(index: 0); //берем лист с которого будем читать данные, будет самый первый лист в файле

                //будем проходить построчно файла Excel и собирать данные, для этого создаем целочисленную переменную
                int rowIndex = 0;

                //используем цикл while, в строке есть данные, она существует, то будем продолжать данный цикл, когда данных не будет цикл закончится
                while (sheet.GetRow(rowIndex) != null)
                {
                    //если в указаной строке в первом столбце GetCell(0) ячейка пустая
                    //либо ячейка пустая во втором столбце GetCell(1), то мы переходим к следующей операции continue
                    if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                        sheet.GetRow(rowIndex).GetCell(1) == null)
                    {
                        //когда мы переходим к следующей операции должны увеличивать rowIndex, поэтому нужно это учитывать, чтобы не попасть в вечный цикл
                        rowIndex++;
                        continue;
                    }
                      

                    //считываем данные из файл Excel: берем лист, берем строку 
                    string name = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                    string number = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;

                    //из всех собранных помещений в моделе ищем помещение с указанным номером
                    var room = rooms.FirstOrDefault(r => r.Number.Equals(number));

                    //ели помещение такое не найдено, то продолжаем цикл
                    if (room == null)
                    {
                        rowIndex++; //здесь также должны увеличивать rowIndex
                        continue;
                    }
                        
                    //если все впорядке то используем конструкцию using и создаем транзакцию в которую записываем данные
                    using(var ts = new Transaction(doc, "Set parameter"))
                    {
                        ts.Start();
                        room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(name);
                        ts.Commit();
                    }

                    rowIndex++; //для того, чтобы переходить каждый раз к новой строке
                }
            }


            return Result.Succeeded;
        }
    }
}
