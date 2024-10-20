namespace SH.ExcelHelper.Treatments
{
    internal class Adjustments
    {
        private readonly SheetHelper _sheetHelper;

        public Adjustments(SheetHelper sheetHelper)
        {
            _sheetHelper = sheetHelper;
        }

        internal string TreatCell(string cellValue, string delimiter)
        {
            //// Header
            //rowFull = table.Columns.Cast<DataColumn>().Select(column =>
            //{
            //    string cellValue = column.ColumnName;
            //    if (cellValue.Contains("\n") || cellValue.Contains("\r")) // Check if the cell contains a line break
            //    {
            //        // Apply double quotes to surround the value and escape the inner double quotes
            //        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
            //    }
            //    return cellValue;
            //}).ToArray();

            //// Row 1
            //rowFull = table.Rows[rowsNumber[0] - 2].ItemArray.Select(cell =>
            //{
            //    string cellValue = cell.ToString();
            //    if (cellValue.Contains("\n") || cellValue.Contains("\r") || cellValue.Contains(delimiter)) // Check if the cell contains a line break
            //    {
            //        // Apply double quotes to surround the value and escape the inner double quotes
            //        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
            //    }
            //    return cellValue;
            //}).ToArray();

            //// Other rows
            //rowFull = table.Rows[rowIndex - 2].ItemArray.Select(cell =>
            //{
            //    string cellValue = cell.ToString();
            //    if (cellValue.Contains("\n") || cellValue.Contains("\r")) // Check if the cell contains a line break
            //    {
            //        // Apply double quotes to surround the value and escape the inner double quotes
            //        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
            //    }
            //    return cellValue;
            //}).ToArray();      

            // Generic
            //return cells.Select(cell =>
            //{
            //    string cellValue = cell.ToString();
            //    if (cellValue.Contains("\n") || cellValue.Contains("\r") || cellValue.Contains(delimiter))
            //    {
            //        cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
            //    }
            //    return (T)Convert.ChangeType(cellValue, typeof(T));
            //}).ToArray();

            if (cellValue.Contains("\n")
                || cellValue.Contains("\r")
                || cellValue.Contains(delimiter)
                || cellValue.Contains("\""))
            {
                cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\""; // Apply ""
            }

            if (_sheetHelper.ProhibitedItems?.Count > 0)
            {
                foreach (var item in _sheetHelper.ProhibitedItems)
                {
                    cellValue = cellValue.Replace(item.Key, item.Value);
                }
            }

            return cellValue;
        }



    }
}
