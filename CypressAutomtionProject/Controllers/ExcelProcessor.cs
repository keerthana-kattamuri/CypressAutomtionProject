using ClosedXML.Excel;
using System.IO;
using System.Text;
using System.Linq;

namespace CypressAutomtionProject.Controllers
{
    //public static class ExcelProcessor
    //{
    //    public static string ProcessExcel(Stream excelStream, string idToFind, string functionTitle)
    //    {
    //        using var workbook = new XLWorkbook(excelStream);
    //        var worksheet = workbook.Worksheets.First();

    //        foreach (var row in worksheet.RowsUsed())
    //        {
    //            foreach (var cell in row.Cells())
    //            {
    //                if (cell.GetValue<string>() == idToFind)
    //                {
    //                    var sb = new StringBuilder();
    //                    sb.Append($"const {functionTitle} = () => {{ return }} ");
    //                    for (int i = 1; i <= worksheet.ColumnCount(); i++)
    //                    {
    //                        string key = worksheet.Cell(1, i).GetValue<string>();
    //                        string value = row.Cell(i).GetValue<string>();
    //                        sb.Append($"{key}: '{value}', ");
    //                    }
    //                    sb.Length -= 2; // Remove the last comma and space
    //                    sb.Append(" }; };");
    //                    return sb.ToString();
    //                }
    //            }
    //        }
    //        return $"// No data found for ID {idToFind}";
    //    }
    //}

    public class ExcelValidator
    {
        public static string GetIdTitlePairs(Stream excelStream)
        //public static List<string> GetIdTitlePairs(Stream excelStream)
        {
            var pairs = new List<string>();
            using var workbook = new XLWorkbook(excelStream);
            var worksheet = workbook.Worksheets.First();

            // Find columns by name
            var headerRow = worksheet.Row(1).Cells().ToList();
            int idCol = headerRow.FindIndex(c => c.GetString().Trim().ToLower() == "id") + 1;
            int titleCol = headerRow.FindIndex(c => c.GetString().Trim().ToLower() == "title") + 1;

            if (idCol == 0 || titleCol == 0)
                throw new Exception("ID or Title column not found.");

            // Loop through data rows
                var sb = new StringBuilder();
            foreach (var row in worksheet.RowsUsed().Skip(1)) // skip header
            {
                var id = row.Cell(idCol).GetString().Trim();
                var title = row.Cell(titleCol).GetString().Trim();

                sb.AppendLine($"it('*{id}* {title}', () => {{ }}); ");

                //if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(title))
                //    pairs.Add($"{id}-{title}");
            }


            return sb.ToString();
        }
    }
}

