using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIReadingDataFromExcelFile
{
    public static class SheetExts
    {
        public static void SetCelValue<T>(this ISheet sheet, int rowIndex, int columnIndex, T value)
        {
            var cellReference = new CellReference(rowIndex, columnIndex); //указываем ссылку на ячейку
            var row = sheet.GetRow(cellReference.Row); //создаем новую переменную

            //если строки с таким индексом cellReference.Row нету, то нужно будет ее создать, таким образом создаем новую строку
            if (row == null)
                row = sheet.GetRow(cellReference.Row);

            //создаем ссылку на конкретную ячейку, по умолчанию ее тоже может не быть
            var cell = row.GetCell(cellReference.Col);

            //если ячейка пустая, то создаем ее
            if (cell == null)
                cell = row.CreateCell(cellReference.Col);

            //нужно проверить, что является значением, которое нужно будет записать в ячейку
            //проверяем: если передан сюда текст, то записываем данное значение как текст
            if (value is string)
            {
                cell.SetCellValue((string)(object)value);//делаем преобразование, по скольку у нас значение обобщенное value, по этому
                                                         //сначала должны преобразовать данное значение в object, а потом уже в string
            }

            //иначе, если значение у нас double,то в этом случае 
            else if (value is double)
            {
                cell.SetCellValue((double)(object)value);
            }

            //если значение целочисленное int
            else if (value is int)
            {
                cell.SetCellValue((int)(object)value);
            }

        }
    }
}
