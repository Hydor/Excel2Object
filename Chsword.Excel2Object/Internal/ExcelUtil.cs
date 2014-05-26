using System.Collections.Generic;
using System.Reflection;

namespace Chsword.Excel2Object.Internal
{
    internal class ExcelUtil
    {
        public static Dictionary<PropertyInfo, ExcelAttribute> GetExportAttrDict<T>()
        {
            var dict = new Dictionary<PropertyInfo, ExcelAttribute>();
            foreach (var propertyInfo in typeof(T).GetProperties())
            {
                var attr = new object();
                var ppi = propertyInfo.GetCustomAttributes(true);
                for (int i = 0; i < ppi.Length; i++)
                {
                    if (ppi[i] is ExcelAttribute)
                    {
                        attr = ppi[i];
                        break;
                    }
                }

                if (attr != null)
                {

                    dict.Add(propertyInfo, attr as ExcelAttribute);

                }
            }
            return dict;
        }
    }
}
