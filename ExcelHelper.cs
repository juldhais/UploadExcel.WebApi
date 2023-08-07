using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace UploadExcel.WebApi;

public static class ExcelHelper
{
    public static List<T> Import<T>(string filePath) where T : new()
    {
        XSSFWorkbook workbook;
        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(stream);
        }

        var sheet = workbook.GetSheetAt(0);

        var rowHeader = sheet.GetRow(0);
        var colIndexList = new Dictionary<string, int>();
        foreach (var cell in rowHeader.Cells)
        {
            var colName = cell.StringCellValue;
            colIndexList.Add(colName, cell.ColumnIndex);
        }

        var listResult = new List<T>();
        var currentRow = 1;
        while (currentRow <= sheet.LastRowNum)
        {
            var row = sheet.GetRow(currentRow);
            if (row == null) break;

            var obj = new T();

            foreach (var property in typeof(T).GetProperties())
            {
                if (!colIndexList.ContainsKey(property.Name))
                    throw new Exception($"Column {property.Name} not found.");

                var colIndex = colIndexList[property.Name];
                var cell = row.GetCell(colIndex);

                if (cell == null)
                {
                    property.SetValue(obj, null);
                }
                else if (property.PropertyType == typeof(string))
                {
                    cell.SetCellType(CellType.String);
                    property.SetValue(obj, cell.StringCellValue);
                }  
                else if (property.PropertyType == typeof(int))
                {
                    cell.SetCellType(CellType.Numeric);
                    property.SetValue(obj, Convert.ToInt32(cell.NumericCellValue));
                }  
                else if (property.PropertyType == typeof(decimal))
                {
                    cell.SetCellType(CellType.Numeric);
                    property.SetValue(obj, Convert.ToDecimal(cell.NumericCellValue));
                }
                else if (property.PropertyType == typeof(DateTime))
                {
                    property.SetValue(obj, cell.DateCellValue);
                }
                else if (property.PropertyType == typeof(bool))
                {
                    cell.SetCellType(CellType.Boolean);
                    property.SetValue(obj, cell.BooleanCellValue);
                }
                else
                {
                    property.SetValue(obj, Convert.ChangeType(cell.StringCellValue, property.PropertyType));
                }  
            }

            listResult.Add(obj);
            currentRow++;
        }

        return listResult;
    }
}
