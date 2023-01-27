using ScriptSolution.Model.Statistics;
using ScriptSolution;
using SourceEts.CommonSettings.Testing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Windows.Media;
using System.Windows.Documents;
using System.Diagnostics.SymbolStore;
using System.Windows.Controls;
using ScriptSolution.Model;
using ScriptSolution.Model.OptimizationResul;
using System.Text.Unicode;

namespace EstimationStatNorm
{
    public class EstimationStatNorm : StatisticScript
    {
        StatContainer? _stat = null;

        /// <summary>
        ///  Начало цикла оптимизации
        /// </summary>
        public override void StartOptimization()
        {
            _stat = new StatContainer();
        }

        /// <summary>
        ///  Окончание каждого прохода оптимизации
        /// </summary>
        //public override void SetUserStatisticParamOnEndTest(Statistic stat)
        public override void SetUserStatisticParamOnEndTest( ReportEndOptimForStatistic report )
        {
            var oParams = OptimParams;
            var uStat = UserStatistics;
            Statistic stat = report.Statistic;

            if( 
                stat.NetProfitLossPercent >= uStat[0].Value  && 
                stat.MaxDrownDownPercent <= uStat[1].Value &&
                stat.FactorRecovery >= uStat[2].Value  &&
                stat.ProfitDealsPercents >= uStat[3].Value
               )
            {
                EstimationStat est = new EstimationStat();

                est.profit = stat.NetProfitLossPercent;
                est.averageDeal = stat.AvaregeProfitLossPercent;
                est.recoveryFactor = stat.FactorRecovery;
                est.dealCount = stat.TradeDealsList.Count;
                est.drawDown = Math.Abs( stat.MaxDrownDownPercent );
                est.symbol = report.Symbol;

                if (report.TimeFrameType.Equals("Минута"))
                    est.tfType = "Minute";
                else if (report.TimeFrameType.Equals("День"))
                    est.tfType = "Day";
                else
                    est.tfType = "Unknown";

                est.tfPeriod = report.TimeFramePeriod;

                if( oParams.Count > 2 )
                {
                    int gebug = 0;
                }

                for (int i=0; i< oParams.Count; i++)
                {
                    UserParam up = new UserParam();
                    up.name = oParams[i].Name;
                    up.value = oParams[i].Value;
                    est.userParam.Add(up);
                }
                _stat.AddStat( est );
            }
        }

        /// <summary>
        ///  Окончание цикла оптимизации
        /// </summary>
        public override void EndOtimizationAll()
        {
            _stat.BuildReport();
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
            DesParamStratetgy.Version = "3";
            DesParamStratetgy.DateRelease = "16.01.2023";
            DesParamStratetgy.DateChange = "24.01.2023";
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "EstimationStatNorm";
        }
    }

    /// <summary>
    ///  СТАТИСТИКА
    /// </summary>
    public class UserParam
    {
        public string name = "";
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
        public string symbol = "";
        public string tfType = "";
        public int tfPeriod = 0;
        public double total = 0; // Общий критерий оценки
        public double normTotal = 0; // Нормированный Общий критерий оценки
        public double paramTotal = 0; // Общий фактор по параметрам
        public double paramNorm = 0; // Нормированный критерий по параметроам
        public List<UserParam>? userParam = null;// Параметры оптимизации

        public EstimationStat()
        {
            userParam = new List<UserParam>();
        }
    };

    /// <summary>
    ///  СТАТИСТИКА
    /// </summary>
    public class StatContainer
    {
        List<List<EstimationStat>> _statContainer;
        string[] _statNames = { "Profit", "DD", "Recovery", "Avg. Deal", "Deal Count" };

        public StatContainer()
        {
            _statContainer = new List<List<EstimationStat>>();
        }

        /// <summary>
        ///  Возвращаем список стат. по ключу
        /// </summary>
        List<EstimationStat> GetStatList(string symbol, string tfType, int tfPeriod)
        {
            if (_statContainer.Count == 0)
                return null;

            for (int i = 0; i < _statContainer.Count; i++)
            {
                if (_statContainer[i].Count > 0)
                {
                    EstimationStat stat = _statContainer[i][0];

                    if (stat.symbol.Equals(symbol) && stat.tfType.Equals(tfType) && stat.tfPeriod == tfPeriod)
                    {
                        return _statContainer[i];
                    }
                }
            }
            return null;
        }

        /// <summary>
        ///  Добавляем статистку по проходу
        /// </summary>
        public void AddStat(EstimationStat stat)
        {
            List<EstimationStat> statList = GetStatList(stat.symbol, stat.tfType, stat.tfPeriod);

            if (statList == null)
            {
                statList = new List<EstimationStat>();
                _statContainer.Add(statList);
            }
            statList.Add(stat);
        }

        /// <summary>
        /// Нормализация параметра
        /// </summary>
        public void BuildReport()
        {
            CaclNormStat();
            CaclNormParams();
            string dir = "C:\\DOCS\\";

            for (int i=0; i<_statContainer.Count; i++)
            {
                string fName = dir + _statContainer[i][0].symbol + "_";
                fName += _statContainer[i][0].tfType + "_";
                fName += _statContainer[i][0].tfPeriod + ".csv";

                SaveResultCSV(_statContainer[i], fName, ";");
            }
        }
        /// <summary>
        /// Нормализация параметра
        /// </summary>
        double CalcNorm(double value, double min, double max)
        {
            double norm = max - min;

            if (norm <= 0)
                return double.NaN;

            double res = (value - min) / norm;

            if (res < 0)
                res = 0;

            if (res > 1)
                res = 1;

            return res;
        }

        /// <summary>
        /// Рсчет нормализованного стат. фактора 
        /// </summary>
        void CaclNormStat()
        {
            for (int k = 0; k < _statContainer.Count; k++)
            {
                List<EstimationStat> estStat = _statContainer[k];

                if (estStat.Count == 0)
                    continue;

                double profMin = estStat.Min(a => a.profit);
                double ddMin = estStat.Min(a => a.drawDown);
                double recMin = estStat.Min(a => a.recoveryFactor);
                double avrMin = estStat.Min(a => a.averageDeal);

                double profMax = estStat.Max(a => a.profit);
                double ddMax = estStat.Max(a => a.drawDown);
                double recMax = estStat.Max(a => a.recoveryFactor);
                double avrMax = estStat.Max(a => a.averageDeal);

                for (int i = 0; i < estStat.Count; i++)
                {
                    estStat[i].total = CalcNorm(estStat[i].profit, profMin, profMax) -
                                       CalcNorm(estStat[i].drawDown, ddMin, ddMax) +
                                       CalcNorm(estStat[i].recoveryFactor, recMin, recMax) +
                                       CalcNorm(estStat[i].averageDeal, avrMin, avrMax);
                }

                double totalMin = estStat.Min(a => a.total);
                double totalMax = estStat.Max(a => a.total);

                for (int i = 0; i < estStat.Count; i++)
                    estStat[i].normTotal = CalcNorm(estStat[i].total, totalMin, totalMax);
            }
        }
        /// <summary>
        /// Рсчет нормализованного стат. фактора 
        /// </summary>
        void CaclNormParams()
        {
            for (int l = 0; l < _statContainer.Count; l++)
            {
                List<EstimationStat> estStat = _statContainer[l];

                if (estStat.Count == 0)
                    continue;

                List<double> min = new List<double>();
                List<double> max = new List<double>();

                for (int i = 0; i < estStat[0].userParam.Count; i++)
                {
                    min.Add(estStat.Min(a => a.userParam[i].value));
                    max.Add(estStat.Max(a => a.userParam[i].value));
                }

                for (int i = 0; i < estStat.Count; i++)
                {
                    estStat[i].paramTotal = 0;

                    for (int k = 0; k < estStat[i].userParam.Count; k++)
                    {
                        estStat[i].paramTotal += CalcNorm(estStat[i].userParam[k].value, min[k], max[k]);
                    }
                }
                double totalMin = estStat.Min(a => a.paramTotal);
                double totalMax = estStat.Max(a => a.paramTotal);

                for (int i = 0; i < estStat.Count; i++)
                    estStat[i].paramNorm = CalcNorm(estStat[i].paramTotal, totalMin, totalMax);
            }
        }


        /// <summary>
        /// Сохранение результатов в CSV
        /// </summary>
        void SaveResultCSV( List<EstimationStat> estStat, string fNameFull, string delimiter)
        {
            if ( estStat.Count <= 0)
                return;

            using (StreamWriter sw = new StreamWriter(fNameFull, false, System.Text.Encoding.Default))
            {
                string str = "";

                for (int i = 0; i < _statNames.Length; i++)
                {
                    str += _statNames[i] + delimiter;
                }

                str += "Symbol" + delimiter;
                str += "Time Type" + delimiter;
                str += "Time" + delimiter;
                str += "Stat Total" + delimiter;
                str += "Stat Norm" + delimiter;

                for (int i = 0; i < estStat[0].userParam.Count; i++)
                {
                    str += estStat[0].userParam[i].name + delimiter;
                }
                str += "Param. Total" + delimiter;
                str += "Param. Norm" + delimiter;
                sw.WriteLine(str);

                for (int i = 0; i < estStat.Count; i++)
                {
                    str = estStat[i].profit.ToString() + delimiter;
                    str += estStat[i].drawDown.ToString() + delimiter;
                    str += estStat[i].recoveryFactor.ToString() + delimiter;
                    str += estStat[i].averageDeal.ToString() + delimiter;
                    str += estStat[i].dealCount.ToString() + delimiter;
                    str += estStat[i].symbol + delimiter;
                    str += estStat[i].tfType + delimiter;
                    str += estStat[i].tfPeriod.ToString() + delimiter;
                    str += estStat[i].total.ToString() + delimiter;
                    str += estStat[i].normTotal.ToString() + delimiter;

                    for (int j = 0; j < estStat[i].userParam.Count; j++)
                    {
                        str += estStat[i].userParam[j].value.ToString() + delimiter;
                    }
                    str += estStat[i].paramTotal.ToString() + delimiter;
                    str += estStat[i].paramNorm.ToString();
                    sw.WriteLine(str);
                }
                sw.Close();
            }
        }
    }
}
