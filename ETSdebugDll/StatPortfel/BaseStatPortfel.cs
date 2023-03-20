using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Interop;
using ScriptSolution;
using ScriptSolution.Model;
using ScriptSolution.Model.Portfels;
using ScriptSolution.Model.Portfels.PortfelTest;
using ScriptSolution.Model.Statistics;
using ScriptSolution.ScanerModel;
using SourceEts.CommonSettings.Testing;

namespace ETSdebugDll.StatPortfel
{
    public class BaseStatPortfel : PortfelScript
    {
        #region Реальная торговля

        public ParamOptimization PausePortfelCalc = new ParamOptimization(60, 1, 10, 5, "Портфель , сек ", "Проверка закрытия позиции, проходит раз в указанное время");

        /// <summary>
        /// Метод для реальной торговли
        /// </summary>
        public override void Execute()
        {

        }

        #endregion


        #region Тестирование портфелей

        public override List<List<int>> ExecuteEndOptimization(List<PortfelResultTest> results, List<CorelationModel> corelations)
        {
            var userParams = UserParams; //пользовательски параметры для скрипта используются только для тестирования
            // brms[0].Number - номер робота, его передавать для 
            return new List<List<int>>();
        }


        public override List<UserParamModel> CreateUserStatisticParam()
        {
            List<UserParamModel> list = new List<UserParamModel>
            {
                new UserParamModel { Name = "Profit ( % ) >= ", Value = 1, Weight = 1, IsOptimization = false },
                new UserParamModel { Name = "Drawdown ( % ) <= ", Value = 50, Weight = 1, IsOptimization = false },
                new UserParamModel { Name = "Recovery Factor >= ", Value = 1, Weight = 1, IsOptimization = false },
                new UserParamModel { Name = "Deal profit ( % ) >= ", Value = 0.3, Weight = 1, IsOptimization = false }
            };
            return list;
        }



        #endregion



        public override void GetAttributesPortfel()
        {
            DesParamStratetgy.Version = "1.0.0.0";
            DesParamStratetgy.DateRelease = "22.04.2019";
            DesParamStratetgy.DateChange = "22.04.2019";
            DesParamStratetgy.Description = "Тестовый портфель описание";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "Отладочный портфель";
        }
    }
}
