using ExcelUtility.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using static ExcelUtility.Attributes.AliasAttribute;

namespace ExcelUtility
{
    public static class PropertyInfoHelper
    {
        public static List<PropertyInfo> GetPropertyInfos<T>()
        {
            return typeof(T)
               .GetProperties()
               //过滤Properties
               .Where(r =>
               {
                   var attr = r.GetCustomAttribute(typeof(DisplayAttribute)) as DisplayAttribute;
                   if (attr != null && !attr.IsDisplay)
                       return false;
                   return true;
               })
               .ToList();
        }

        /// <summary>
        /// 获取类型中的所有属性包括列表中的属性（DisplayAttribute = false的属性不显示，属性中 AliasAttributeType = List 的集合的属性才添加）
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static List<PropertyInfo> GetPropertyInfosWithListType(Type type)
        {
            List<PropertyInfo> PropertyInfos = new List<PropertyInfo>();
            foreach (var item in GetPropertyInfos(type))
            {
                if (GetAliasType(item) == AliasAttributeType.List)
                {
                    PropertyInfos.AddRange(GetPropertyInfosWithListType(item.PropertyType.GenericTypeArguments[0]));
                }
                else
                {
                    PropertyInfos.Add(item);
                }
            }
            return PropertyInfos;
        }

        public static List<PropertyInfo> GetPropertyInfos(Type type)
        {
            return type
               .GetProperties()
               //过滤Properties
               .Where(r =>
               {
                   var attr = r.GetCustomAttribute(typeof(DisplayAttribute)) as DisplayAttribute;
                   if (attr != null && !attr.IsDisplay)
                       return false;
                   return true;
               })
               .ToList();
        }

        public static string GetAliasName(PropertyInfo pi) => (pi.GetCustomAttribute(typeof(AliasAttribute)) as AliasAttribute)?.Alias ?? pi.Name;

        public static AliasAttributeType GetAliasType(PropertyInfo pi) => (pi.GetCustomAttribute(typeof(AliasAttribute)) as AliasAttribute)?.Type ?? default(AliasAttributeType);

        public static string GetAlias(Type type) => (type.GetCustomAttributes(typeof(AliasAttribute), false).FirstOrDefault() as AliasAttribute)?.Alias ?? type.Name;

        public static bool ObjectIsAllNullOrEmpty<T>(T t)
        {
            if (t == null)
                return true;
            var c = t.GetType().GetProperties().All(s => s.GetValue(t) == null || s.GetValue(t).ToString() == "");
            return c;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="propertyInfo">obj 对应的属性</param>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static object GetDefaultOrValueByConvertAttribute(PropertyInfo propertyInfo, object obj)
        {
            if (obj == null)
                return null;
            if (propertyInfo == null)
                return obj;
            var attr = propertyInfo.GetCustomAttribute(typeof(BaseConvertAttribute)) as BaseConvertAttribute;
            return attr == null ? obj : attr.GetValue(obj);
        }
    }
}