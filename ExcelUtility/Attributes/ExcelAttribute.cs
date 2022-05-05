using System;

namespace ExcelUtility.Attributes
{
    public enum AliasAttributeType
    {
        Default,
        List,
    }

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Class, AllowMultiple = false)]
    public class AliasAttribute : Attribute
    {
        /// <summary>
        /// 名字
        /// </summary>
        public string Alias { get; set; }

        /// <summary>
        /// 类型
        /// </summary>
        public AliasAttributeType Type { get; set; }

        public abstract class BaseConvertAttribute : Attribute
        {
            public abstract object GetValue(object obj);

            public abstract Type GetType();
        }
    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class DisplayAttribute : Attribute
    {
        public bool IsDisplay { get; set; }
    }
}