using ExcelUtility.Model.Base;
using ExcelUtility.Simple.Model;
using System.Collections.Generic;

namespace ExcelUtility.Simple.Tables
{
    public class UserTable : BaseExcelSheet<User>
    {
        public UserTable()
        {
            SheetStyle.ShowRowNum = true;
            SheetStyle.ChartInfo = new ChartInfo(1, 1, 9, 18)
            {
                CategoryColumnIndex = 1,
                ValueColumnIndex = 0,
                SerieTitle = "采集电压",
                CatAxisTitle = "测试时间",
                ValueAxisTitle = "采集电压",
                ChartTitle = "采集电压 \\ 测试时间(记录)"
            };
            Datas = new List<User>();
            for (int i = 0; i < 100; i++)
            {
                Datas.Add(new User()
                {
                    Name = $"小明{i}",
                    Age = i,
                    Works = new List<Work>() { new Work() { Name = $"任务{i}", Time = i.ToString() }, new Work() { Name = $"任务{i + 1}", Time = (i + 1).ToString() }, new Work() { Name = $"任务{i + 2}", Time = (i + 2).ToString() } },
                    EXWorks = new List<Work>() { new Work() { Name = $"任务{i}", Time = i.ToString() }, new Work() { Name = $"任务{i + 1}", Time = (i + 1).ToString() }, new Work() { Name = $"任务{i + 2}", Time = (i + 2).ToString() } },
                });
            }
        }
    }
}