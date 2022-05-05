using ExcelUtility.Attributes;

namespace ExcelUtility.Simple.Model
{
    public class Work
    {
        [Alias(Alias = "任务名称")]
        public string Name { get; set; }

        [Display(IsDisplay = false)]
        public string Time { get; set; }
    }
}