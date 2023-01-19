using ScriptSolution.Model.Statistics;
using ScriptSolution;
using SourceEts.CommonSettings.Testing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Data;
using System.Windows.Media;
using System.Windows.Documents;
using System.Diagnostics.SymbolStore;
using System.Windows.Controls;
using ScriptSolution.Model;

namespace EstimationStatNorm
{
    public class UserParam
    {
        public string  name = "";
        public double value = 0;
        public int digits = 6;
    }
    public class EstimationStat
    {
        public double profit = 0; // Профит
        public double drawDown = 0; // Максимальная просадка
        public double recoveryFactor = 0; // Факттор восстановления
        public double averageDeal = 0; // Средняя прибыль на сделку
        public double dealCount = 0; // Количество сделок 
        public double total = 0; // Общий критерий оценки
        public double normTotal = 0; // Нормированный Общий критерий оценки
        public List<UserParam> userParam;// Параметры оптимизации
    };

    public class StatisticDebugScript : StatisticScript
    {
        int _passCount = 0;
        List<EstimationStat> ? _estStat = null;
        string[] _statNames = { "Profit", "DD", "Recovery", "Avg. Deal", "Deal Count" };

        /// <summary>
        ///  Начало цикла оптимизации
        /// </summary>
        public override void StartOptimization()
        {
            if (_estStat == null)
                _estStat = new List<EstimationStat>();
            else
                _estStat.Clear();
            _passCount = 0;
        }

        /// <summary>
        ///  Окончание прохода оптимизации
        /// </summary>
        public override void SetUserStatisticParamOnEndTest(Statistic stat)
        {
            var oParams = OptimParams;
            var uStat = UserStatistics;

            if (
                 ( uStat[0].Value < 0 || stat.NetProfitLossPercent >= uStat[0].Value ) &&
                 ( uStat[1].Value < 0 || stat.MaxDrownDownPercent <= uStat[1].Value) &&
                 ( uStat[2].Value < 0 || stat.FactorRecovery >= uStat[2].Value ) &&
                 ( uStat[3].Value < 0 || stat.ProfitDealsPercents >= uStat[3].Value)
               )
            {
                EstimationStat est = new EstimationStat();
                est.userParam = new List<UserParam>();

                est.profit = stat.NetProfitLossPercent;
                est.averageDeal = stat.ProfitDealsPercents;
                est.recoveryFactor = stat.FactorRecovery;
                est.dealCount = stat.TradeDealsList.Count;
                est.drawDown = stat.MaxDrownDownPercent;
                
                for(int i=0; i< oParams.Count; i++)
                {
                    UserParam up = new UserParam();
                    up.name = oParams[i].Name;
                    up.value = oParams[i].Value;
                    est.userParam.Add(up);
                }
                _estStat.Add(est);
            }
            _passCount++;
        }
        /// <summary>
        /// Рсчет нормализованного фактора оценки
        /// </summary>
        
        /// <summary>
        ///  Окончание цикла оптимизации
        /// </summary>
        public override void EndOtimizationAll()
        {
            SaveResultCSV("C:\\DOCS\\out.csv", ";");
        }

        /// <summary>
        /// Сохранение результатов в CSV
        /// </summary>
        void SaveResultCSV(string fNameFull, string delimiter )
        {
            if (_estStat.Count <= 0)
                return;

            using (StreamWriter sw = new StreamWriter(fNameFull, false, System.Text.Encoding.Default))
            {
                string str = "";

                for( int i = 0; i < _statNames.Length; i++ )
                {
                    str += _statNames[i] + delimiter;
                }

                for( int i=0; i<_estStat[0].userParam.Count; i++ )
                {
                    str += _estStat[0].userParam[i].name;

                    if (i < _estStat[0].userParam.Count - 1)
                        str += delimiter;
                }

                sw.WriteLine(str);

                for(int i=0; i<_estStat.Count; i++)
                {
                    str = _estStat[i].profit.ToString() + delimiter;
                    str += _estStat[i].drawDown.ToString() + delimiter;
                    str += _estStat[i].recoveryFactor.ToString() + delimiter;
                    str += _estStat[i].averageDeal.ToString() + delimiter;
                    str += _estStat[i].dealCount.ToString() + delimiter; 

                    for( int j=0; j< _estStat[i].userParam.Count; j++)
                    {
                        str += _estStat[i].userParam[j].value.ToString() ;

                        if ( j < _estStat[i].userParam.Count - 1)
                            str += delimiter;
                    }
                    sw.WriteLine(str);
                }
                sw.Close();
            }
        }

        /// <summary>
        /// Определение параметров для скрипта
        /// </summary>
        public override List<UserParamModel> CreateUserStatisticParam()
        {
            List<UserParamModel> list = new List<UserParamModel>();
            list.Add(new UserParamModel { Name = "Profit ( % ) >= ", IsOptimization = false });
            list.Add(new UserParamModel { Name = "Drawdown ( % ) <= ", IsOptimization = false });
            list.Add(new UserParamModel { Name = "Recovery Factor >= ", IsOptimization = false });
            list.Add(new UserParamModel { Name = "Deal profit ( % ) >= ", IsOptimization = false });
            return list;
        }

        public override void GetAttributes()
        {
            DesParamStratetgy.Version = "1";
            DesParamStratetgy.DateRelease = "16.01.2023";
            DesParamStratetgy.DateChange = "17.01.2023";
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "EstimationStatNorm";
        }
    }
}
