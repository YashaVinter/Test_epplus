using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;
namespace Test_epplus
{
    class Program
    {
        static ExcelRangeBase Find(in ExcelRange obj, in ExcelRange range) {
            
            var list_range = range.ToList();
            var obj_value = Convert.ToInt32(obj.Value);
            for (int i = 0; i < list_range.Count; ++i) {
                var tmp = Convert.ToInt32(list_range.ElementAt(i).Value);
                if (tmp == obj_value) {
                    return list_range.ElementAt(i);
                }
            }
            return null;
        }

        static ExcelRangeBase Find(in ExcelRangeBase obj, in ExcelRange range)
        {

            var list_range = range.ToList();
            var obj_value = Convert.ToInt32(obj.Value);
            for (int i = 0; i < list_range.Count; ++i)
            {
                var tmp = Convert.ToInt32(list_range.ElementAt(i).Value);
                if (tmp == obj_value)
                {
                    return list_range.ElementAt(i);
                }
            }
            return null;
        }

        static void MakeRef(in ExcelRangeBase source, in ExcelRangeBase destination)
        {
            if (source.Count() != destination.Count())
                throw new Exception();
            var cell_destination = destination.First();
            //destination.First().Current;
            foreach (var cell_source in source)
            {
                var formula = cell_source.FullAddress;
                cell_destination.Formula = formula;
                destination.MoveNext();
                cell_destination = destination.Current;
            }
        }



        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var path = new FileInfo(@"C:\Users\User\Desktop\Работа Энергоком\2021 Сургут\Протоколы+текучка\Наладка\Комплексное опробование\Прогрузка токовых цепей\09,11,2021\Копия Копия Данные 08,11,2021 в2 тест.xlsm");
            using (ExcelPackage excel = new ExcelPackage(path))
            {
                // Target a worksheet
                var sheet_trans = excel.Workbook.Worksheets["По трансформаторам 2"];
                var sheet_terminal = excel.Workbook.Worksheets["По терминалам 2"];
                excel.Workbook.CalcMode = ExcelCalcMode.Automatic;

                foreach (var cell_source in sheet_terminal.Cells["B11:B89"]) {
                    //var cell_source = sheet_terminal.Cells["B11"];
                    var range = sheet_trans.Cells["B11:B89"];
                    var cell_destination = Find(in cell_source, in range);

                    var obj_s = cell_source.Offset(0, 3);
                    var obj_d = cell_destination.Offset(0, 3);
                    MakeRef(in obj_s, in obj_d);

                    obj_s = sheet_terminal.Cells[cell_source.Offset(0, 5).FullAddress + ":" + cell_source.Offset(0, 15).FullAddress]; // 05 07 did
                    obj_d = sheet_trans.Cells[cell_destination.Offset(0, 5).FullAddress + ":" + cell_destination.Offset(0, 15).FullAddress];
                    MakeRef(in obj_s, in obj_d);
                    cell_source.Offset(0, -1).Value = 1;
                    cell_destination.Offset(0, -1).Value = 1;
                }
                excel.SaveAs(path); 
            }
        }
    }
}
/*
 1. лучше сделать макереф не ссылочный
2. сделать функцию для obj_s и obj_d, чтобы в нее диапазон передавать
3. учесть отсутствие ячеек при поиске, пример 93
 */