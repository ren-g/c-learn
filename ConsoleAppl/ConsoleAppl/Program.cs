using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using ConsoleAppl.Global;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;

namespace ConsoleAppl
{
    class Program
    {
        static void Main()
        {
            string sql = @"select Sex, Register_ID, Nation, NativePlace, IdCardNo from Pati_Regi_BasicInfo ";
            DataSet ds = DB.getDataSet(sql);
            List<PresenInfo> preList = TableInfo<PresenInfo>.GetList(ds.Tables[0]).ToList();
            ExcelInfoCallback getP = new ExcelInfoCallback(new Program().GetTrueValue);

            string sql2 = @"select * from Pati_Regi_Diagnos ";
            DataSet ds2 = DB.getDataSet(sql2);
            List<PresenInfo> preList2 = TableInfo<PresenInfo>.GetList(ds2.Tables[0]).ToList();

            IWorkbook wb = new HSSFWorkbook();
            OutPutExcel<PresenInfo>.ListToExcel(wb, preList, "sdd", getP);
            using (FileStream stm = File.Create("E:\\Desktop\\1.xls"))
            {
                wb.Write(stm);
            }
        }        
        public string GetTrueValue (string name,object obj){
            if (name == "Sex")
            {
                string str = "";
                switch (obj.ToString())
                {
                    case "1": str = "男"; break;
                    case "2": str = "女"; break;
                }
                return str;
            }
            return obj.ToString();
        }
    }
}
