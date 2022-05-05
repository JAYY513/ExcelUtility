using ExcelUtility.Attributes;
using System.Collections.Generic;

namespace ExcelUtility.Simple.Model
{
    public class User
    {
        [Alias(Alias = "名称")]
        public string Name { get; set; }

        [Display(IsDisplay = false)]
        public int Age { get; set; }

        [Alias(Type = AliasAttributeType.List, Alias = "任务列表")]
        [Display(IsDisplay = true)]
        public List<Work> Works { get; set; }

        [Display(IsDisplay = false)]
        public List<Work> EXWorks { get; set; }
    }
}