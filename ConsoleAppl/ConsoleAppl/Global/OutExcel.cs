using System.Collections.Generic;
using System.Reflection;
using NPOI.SS.UserModel;
using System.IO;
using System.Data;
using System.Text;
using NPOI.HSSF.UserModel;
namespace ConsoleAppl.Global
{
    public delegate string ExcelInfoCallback(string key, object oldvalue);
    public class OutPutExcel<T> where T:new()
    {
        public static void DtatsetToExcel(DataSet ds, ExcelInfoCallback callack = null) 
        {
            IWorkbook wb = new HSSFWorkbook();
            StringBuilder str = new StringBuilder();
            foreach (DataTable dt in ds.Tables) 
            {
                ISheet sheet = wb.CreateSheet(dt.TableName);
                IRow RowHeader = sheet.CreateRow(0);//表头
                for (int col = 0; col < dt.Columns.Count; col++) 
                {
                    RowHeader.CreateCell(col).SetCellValue(dt.Columns[col].ColumnName);
                }
                string name, ExcelStr;
                for (int rows = 0; rows < dt.Rows.Count; rows++) 
                {
                    IRow Row = sheet.CreateRow(rows + 1);
                    for (int cols = 0; cols < dt.Columns.Count; cols++) 
                    {
                        name = dt.Columns[cols].ColumnName;
                        object obj = dt.Rows[rows][cols];
                        if (obj == null || obj.ToString() == "")
                        {
                            ExcelStr = "空";
                        }
                        else {
                            if (callack != null)
                            {
                                ExcelStr = callack(name, obj);
                            }
                            else {
                                ExcelStr = obj.ToString();
                            }
                        }
                        Row.CreateCell(cols).SetCellValue(ExcelStr);
                    }
                }
            }
            using (FileStream stm = File.Create("C:\\Users\\rjl\\Desktop\\1.xls"))
            {
                wb.Write(stm);
            }
        }
        #region 泛型数组遍历
        public static void ListToExcel(IWorkbook wb, List<T> list, string tableName = null, ExcelInfoCallback callack = null)
        {
            string name, ExcelStr;
            PropertyInfo[] propertys = new T().GetType().GetProperties();          
            ISheet sheet = wb.CreateSheet(tableName);
            IRow rowHeader = sheet.CreateRow(0);//表头
            for (int i=0;i<propertys.Length;i++) {
                rowHeader.CreateCell(i).SetCellValue(propertys[i].Name);
            }
            for (int rows = 0; rows < list.Count; rows++)
            {
                IRow row = sheet.CreateRow(rows + 1);
                for (int cols = 0; cols < propertys.Length; cols++)
                {
                    name = propertys[cols].Name;
                    object obj = list[rows].GetType().GetProperty(name).GetValue(list[rows], null);
                    if (obj == null || obj.ToString() == "")
                    {
                        ExcelStr = "空";
                    }
                    else
                    {
                        if (callack != null)
                        {
                            ExcelStr = callack(name, obj);
                        }
                        else
                        {
                            ExcelStr = obj.ToString();
                        }
                    }
                    row.CreateCell(cols).SetCellValue(ExcelStr);
                }
            }            
        }
        #endregion
    }
}
