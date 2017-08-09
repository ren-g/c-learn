using System.Collections.Generic;
using System.Reflection;
using NPOI.SS.UserModel;
using System.IO;

namespace ConsoleAppl.Global
{
    public delegate string ExcelInfoCallback(string key, object oldvalue);
    public class OutPutExcel<T> where T:new()
    {
        #region 泛型数组遍历
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list">将导出列表</param>
        /// <param name="path">完整路径</param>
        /// <param name="tableName">表名</param>
        /// <param name="callack">过程中需要单独处理的数据</param>
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
