namespace SH.ExcelHelper.Tools
{
    interface ISheetHelper
    {
        void CloseExcel();
        int GetIndexColumn(string columnName);
        string GetNameColumn(int columnIndex);





    }
}
