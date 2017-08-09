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
            string sql = @"select Sex, Register_ID, Nation, NativePlace, IdCardNo from Pati_Regi_BasicInfo;select * from Pati_Regi_Diagnos ";
            DataSet ds = DB.getDataSet(sql);
            ExcelInfoCallback getP = new ExcelInfoCallback(new Program().GetTrueValue);
            OutPutExcel<PresenInfo>.DtatsetToExcel(ds,getP);            
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
