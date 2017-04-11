using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
namespace ExcelConverter
{
    //=========================================================================
    // *  作者：   杨泉耀
    // *  时间：   2016年11月29日14:03:15
    // *  文件名： Reflection      
    // *  版本：   1.0
    // *  说明：   使用反射这一高级功能再加上泛型，创造一个可以创造任何可创造的实体^_^  
    //=========================================================================
    public class Reflection
    {
        public static T CreateInstance<T>(object[] parameter)
        {
            var type = typeof(T);
            var cons = type.GetConstructors(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);
            if (cons == null)
                return default(T);
            if (parameter != null)
            {
                var result = from con in cons where con.GetParameters().Length == parameter.Length select con;
                if (result.Count() > 0)
                {
                    ConstructorInfo targetCon = null;
                    foreach (var pCon in result)
                    {
                        var parameters = pCon.GetParameters();
                        var flag = true;
                        for (int i = 0; i < parameters.Length; i++)
                        {
                            if (parameters[i].ParameterType != parameter[i].GetType())
                            {
                                flag = false;
                                break;
                            }
                        }
                        if (flag)
                        {
                            targetCon = pCon;
                            var targetObject = targetCon.Invoke(parameter);
                            return (T)targetObject;
                        }
                    }
                    return default(T);
                }
                else
                    return default(T);
            }
            else
            {
                var targetCon = (from con in cons where con.GetParameters() == null || con.GetParameters().Length == 0 select con).First();
                var targetObject = targetCon.Invoke(parameter);
                return (T)targetObject;
            }
        }
        public static PropertyInfo[] GetProperties(object obj)
        {
            var type = obj.GetType();
            return type.GetProperties();
        }
        public static PropertyInfo[] GetProperties<T>()
        {
            var type = typeof(T);
            return type.GetProperties();
        }
        public static PropertyInfo[] GetProperties<T>(BindingFlags flage)
        {
            var type = typeof(T);
            return type.GetProperties(flage);
        }
        public static bool ContainsAttribute(PropertyInfo property, Type type)
        {
            var attrs = property.GetCustomAttributes(true);

            foreach (var attr in attrs)
            {
                if (attr.GetType().Equals(type))
                {
                    return true;
                }
            }
            return false;
        }
        public static object GetAttribute(PropertyInfo property, Type type)
        {
            var attrs = property.GetCustomAttributes(true);

            foreach (var attr in attrs)
            {
                if (attr.GetType().Equals(type))
                {
                    return attr;
                }
            }
            return null;
        }
        public static void SetPropertyValue(object context, PropertyInfo property, object value)
        {
            if (value == null)
                return;
            if (property.PropertyType.Equals(value.GetType()))
            {
                property.SetValue(context, value, null);
                return;
            }
            switch (property.PropertyType.FullName)
            { 

                case "System.Byte":
                case "System.UInt16":
                case "System.Int16":
                case "System.Int32":
                case "System.Int64":
                   property.SetValue(context, long.Parse(value.ToString()), null);
                    break;
                case "System.Double":
                case "System.Single":
                    property.SetValue(context, double.Parse(value.ToString()), null);   
                    break;
                case "System.Decimal":
                    property.SetValue(context,Decimal.Parse(value.ToString()), null);
                    break;
                case "System.DateTime":
                    property.SetValue(context, DateTime.Parse(value.ToString()), null);
                    break;
                case "System.Boolean":
                    property.SetValue(context, bool.Parse(value.ToString()), null);
                    break;
                default :
                    property.SetValue(context, value, null);
                    break;
            }
        }


    }

}
