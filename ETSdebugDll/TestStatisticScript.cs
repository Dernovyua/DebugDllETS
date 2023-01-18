using ScriptSolution.Model.Statistics;
using ScriptSolution;
using SourceEts.CommonSettings.Testing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ETSdebugDll
{
    public class StatisticDebugScript : StatisticScript
    {

        public override void SetUserStatisticParamOnEndTest(Statistic stat)
        {
            int countProfit = 0;
            int countLoss = 0;

            for (int i = 0; i < stat.TradeDealsList.Count; i++)
            {
                if (stat.TradeDealsList[i].ProfitLoss > 0)
                    countProfit += 1;
                else
                    countLoss += 1;
            }

            SetUserStatisticParam("Всего +", countProfit);
            SetUserStatisticParam("Всего -", countLoss);
        }



        public override List<UserParamModel> CreateUserStatisticParam()
        {
            List<UserParamModel> list = new List<UserParamModel>();
            list.Add(new UserParamModel { Name = "Всего +", IsOptimization = false });
            list.Add(new UserParamModel { Name = "Всего -", IsOptimization = true });

            return list;
        }

        public override void GetAttributes()
        {
            DesParamStratetgy.Version = "1";
            DesParamStratetgy.DateRelease = "16.01.2023";
            DesParamStratetgy.DateChange = "16.01.2023";
            DesParamStratetgy.Description = "вариант 1, смотрим";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "DllDebug статистика";
        }
    }
}
