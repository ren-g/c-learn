using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace ConsoleAppl.Global
{
    public class TableInfo<T> where T : new()
    {
        #region DataTale转List
        public static IList<T> GetList(DataTable dt) { 
            //定义集合
            IList<T> ts = new List<T>();
            //获得模型类型
            Type type = typeof(T);
            string tempName = "";            
            foreach(DataRow dr in dt.Rows){
                T t = new T();
                PropertyInfo[] propertys = t.GetType().GetProperties();
                foreach (PropertyInfo pi in propertys)
                {
                    //检查dataTable是否包含此列
                    tempName = pi.Name;
                    if (dt.Columns.Contains(tempName)) 
                    {
                        if (!pi.CanWrite) continue;
                        object value = dr[tempName];
                        if (value != DBNull.Value) pi.SetValue(t, value, null);
                    }
                }
                ts.Add(t);
            }
            return ts;
        }
        #endregion
    }
}
