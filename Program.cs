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

        //static void MakeRef(in ExcelRange source, in ExcelRange destination) {
        //    if (source.Count() != destination.Count())
        //        throw new Exception();

        //    var cell_destination = destination.First();
        //    //destination.First().Current;
        //    foreach (var cell_source in source) {
        //        var formula = cell_source.FullAddress;
        //        cell_destination.Formula = formula;
        //        destination.MoveNext();
        //        cell_destination = destination.Current;
        //    }
        //}

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

            //  1 version ( @"C:\Users\User\Desktop\Работа Энергоком\Сургут\Протоколы+текучка\Наладка\Комплексное опробование\Прогрузка токовых цепей\Данные_тест.xlsm")
            //  2 version (@"C:\Users\User\Desktop\Работа Энергоком\Сургут\Протоколы+текучка\Наладка\Комплексное опробование\Прогрузка токовых цепей\Копия Данные 08,11,2021 тест.xlsm")
            var path = new FileInfo(@"C:\Users\User\Desktop\Работа Энергоком\Сургут\Протоколы+текучка\Наладка\Комплексное опробование\Прогрузка токовых цепей\09,11,2021\Копия Копия Данные 08,11,2021 в2 тест.xlsm"); // 
            using (ExcelPackage excel = new ExcelPackage(path))
            {
                // Target a worksheet
                var sheet_trans = excel.Workbook.Worksheets["По трансформаторам 2"];
                var sheet_terminal = excel.Workbook.Worksheets["По терминалам 2"];
                excel.Workbook.CalcMode = ExcelCalcMode.Automatic;
                //sheet_trans.Workbook.CalcMode = ExcelCalcMode.Automatic;
                //sheet_terminal.Workbook.CalcMode = ExcelCalcMode.Automatic;

                /*
                var a = sheet_terminal.Cells["B11"];
                var b = sheet_terminal.Cells["B11"].Offset(0,0);
                ExcelRange gans1488 = (ExcelRange)b;
                */
                //var cells = sheet_terminal.Cells["A3:A4"];
                //object arr = sheet_terminal.Cells["A3:A4"].Value;
                //int a10 = Convert.ToInt32(sheet_terminal.Cells["A3"].Value);
                //var a13 = Find(in obj,in range); // in - const ref
                //var a14 = sheet_terminal.Cells["B11"];

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
                var a1 = 1;
                /*
                 1. лучше сделать макереф не ссылочный
                2. сделать функцию для obj_s и obj_d, чтобы в нее диапазон передавать
                3. учесть отсутствие ячеек при поиске, пример 93
                 */





                //var a4 = cells.ToArray();
                //object a5 = a4.GetValue(0);
                //int[] ar = new int[2];
                //a5.Equals(a5);
                //var a8 = (a5.ToString());
                //a4.CopyTo(ar, 0);
                //var a6 = a4.Length;
                //var a2 = cells.Rows;
                //var a3 = cells.Columns;
                //Type t = arr.GetType();
                //var a1 = t.GetArrayRank();
                //var b = cells.ToList();
                ////b.Find(;
                ////Predicate<ExcelRangeBase> pre = new Predicate<ExcelRangeBase> (;
                //List<int> test = new List<int>();
                ////test.Add(sheet_terminal.Cells["A3"].Value);
                //var seq = sheet_terminal.Cells["A3:A4"].ToList();



                /*
                 sheet_termi)al.Cells["A2"].Value = 101;
                //sheet_terminal.Cells["A2"].Formula = " A1";
                //sheet_terminal.Cells["A2"].Formula = " ПОИСКПОЗ(41;B11:B100;0)"; // =ПОИСКПОЗ(41;B1:B300;0)
                //sheet_terminal.Cells["A2"].Calculate();
                // нужно сделать коллекцию их 
                // excel.Workbook.VbaProject.Modules.Sum()
                */

                excel.SaveAs(path);
                
            }
        }
    }
}

/*
                sheet_terminal.Cells["A5"].Formula = " =SUM(A3:A4)";
                sheet_terminal.Cells["A6"].Formula = " =ПОИСКПОЗ(41;B11:B100;0)"; // not work
                sheet_terminal.Cells["A7"].Formula = " =ПОИСКПОЗ(41,B11:B100,0)";
                sheet_terminal.Cells["A8"].Formula = " ПОИСКПОЗ(41;B11:B100;0)"; // not work
                sheet_terminal.Cells["A9"].Formula = " ПОИСКПОЗ(41,B11:B100,0)";


                 foreach (var cell in sheet_terminal.Cells["B11:B12"]) {
                    if (cell.Value.Equals(42)) {
                        sheet_terminal.Cells["C1"].Value = 3;
                        break;
                    }
                }

                var query =
                    from cell in sheet_terminal.Cells["B11:B12"]
                        // where cell.Value?.ToString() == "CRÉDITOS"
                    where cell.Value.Equals(41) 
                    select cell;
                sheet_terminal.Cells["C1"].Value = query; 
 */

/* сделать новый файл

namespace Test_epplus
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");

                var headerRow = new List<string[]>()
                {
                 new string[] { "ID", "First Name", "Last Name", "DOB" }
                 };

                // Determine the header range (e.g. A1:D1)
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                // Popular header row data
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                var tmp = "A1";
                worksheet.Cells["A2"].Value = 1;
                worksheet.Cells["A2"].Offset(1, 1).Value = 10;

                var path1 = @"C:\Users\User\Desktop\test1.xlsx";
                var path2 = @"C:\Users\User\Desktop\test2.xlsx";
                excel.SaveAs(path1);
                //
                FileInfo existingFile = new FileInfo(path1);
                ExcelPackage excel2 = new ExcelPackage(path2);
                var formula = "A1";
                excel2.Workbook.Worksheets["Worksheet1"].Cells["A3"].Formula = formula;
                excel2.Workbook.Worksheets["Worksheet1"].Cells["A3"].Calculate();
                excel2.Save();
                
            }
        }
    }
}
*/