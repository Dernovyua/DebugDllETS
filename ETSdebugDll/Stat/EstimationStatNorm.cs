using ScriptSolution.Model.Statistics;
using ScriptSolution;
using SourceEts.CommonSettings.Testing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using iTextSharp;
//using iTextSharp.text;
//using iTextSharp.text.pdf;
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
        public string symbol = "";
        public string tfType = "";
        public int tfPeriod = 0;
        public double total = 0; // Общий критерий оценки
        public double normTotal = 0; // Нормированный Общий критерий оценки
        public double paramTotal = 0; // Общий фактор по параметрам
        public double paramNorm = 0; // Нормированный критерий по параметроам
        public List<UserParam> userParam;// Параметры оптимизации
    };

    public class EstimationStatNorm : StatisticScript
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
        //public override void SetUserStatisticParamOnEndTest(Statistic stat)
        public override void SetUserStatisticParamOnEndTest( ReportEndOptimForStatistic report )
        {
            var oParams = OptimParams;
            var uStat = UserStatistics;
            Statistic stat = report.Statistic;

            /*if(
                 ( uStat[0].Value < 0 || stat.NetProfitLossPercent >= uStat[0].Value ) &&
                 ( uStat[1].Value < 0 || stat.MaxDrownDownPercent <= uStat[1].Value) &&
                 ( uStat[2].Value < 0 || stat.FactorRecovery >= uStat[2].Value ) &&
                 ( uStat[3].Value < 0 || stat.ProfitDealsPercents >= uStat[3].Value)
               )*/
            {
                EstimationStat est = new EstimationStat();
                est.userParam = new List<UserParam>();

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

                for (int i=0; i< oParams.Count; i++)
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
        ///  Окончание цикла оптимизации
        /// </summary>
        public override void EndOtimizationAll()
        {
            CaclNormStat();
            CaclNormParams();
            SaveResultCSV("C:\\DOCS\\out.csv", ";");
        }

        /// <summary>
        /// Нормализация оценки
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
            double profMin = _estStat.Min(a => a.profit);
            double ddMin = _estStat.Min(a => a.drawDown);
            double recMin = _estStat.Min(a => a.recoveryFactor);
            double avrMin = _estStat.Min(a => a.averageDeal);

            double profMax = _estStat.Max(a => a.profit);
            double ddMax = _estStat.Max(a => a.drawDown);
            double recMax = _estStat.Max(a => a.recoveryFactor);
            double avrMax = _estStat.Max(a => a.averageDeal);

            for (int i = 0; i < _estStat.Count; i++)
            {
                _estStat[i].total = CalcNorm(_estStat[i].profit, profMin, profMax) -
                                    CalcNorm(_estStat[i].drawDown, ddMin, ddMax) +
                                    CalcNorm(_estStat[i].recoveryFactor, recMin, recMax) +
                                    CalcNorm(_estStat[i].averageDeal, avrMin, avrMax);
            }

            double totalMin = _estStat.Min(a => a.total);
            double totalMax = _estStat.Max(a => a.total);

            for (int i = 0; i < _estStat.Count; i++)
                _estStat[i].normTotal = CalcNorm(_estStat[i].total, totalMin, totalMax);

        }
        /// <summary>
        /// Рсчет нормализованного стат. фактора 
        /// </summary>
        void CaclNormParams()
        {
            List<double> min = new List<double>();
            List<double> max = new List<double>();

            for( int i=0; i < _estStat[0].userParam.Count; i++)
            {
                min.Add(_estStat.Min(a => a.userParam[i].value));
                max.Add(_estStat.Max(a => a.userParam[i].value));
            }

            for (int i = 0; i < _estStat.Count; i++)
            {
                _estStat[i].paramTotal = 0;

                for ( int k=0; k< _estStat[i].userParam.Count; k++)
                {
                    _estStat[i].paramTotal += CalcNorm(_estStat[i].userParam[k].value, min[k], max[k]);
                }
            }
            double totalMin = _estStat.Min(a => a.paramTotal);
            double totalMax = _estStat.Max(a => a.paramTotal);

            for (int i = 0; i < _estStat.Count; i++)
                _estStat[i].paramNorm = CalcNorm(_estStat[i].paramTotal, totalMin, totalMax);
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

                str += "Symbol" + delimiter;
                str += "Time Type" + delimiter;
                str += "Time" + delimiter;
                str += "Stat Total" + delimiter;
                str += "Stat Norm" + delimiter;

                for ( int i=0; i<_estStat[0].userParam.Count; i++ )
                {
                    str += _estStat[0].userParam[i].name + delimiter;
                }
                str += "Param. Total" + delimiter;
                str += "Param. Norm" + delimiter;
                sw.WriteLine(str);

                for(int i=0; i<_estStat.Count; i++)
                {
                    str = _estStat[i].profit.ToString() + delimiter;
                    str += _estStat[i].drawDown.ToString() + delimiter;
                    str += _estStat[i].recoveryFactor.ToString() + delimiter;
                    str += _estStat[i].averageDeal.ToString() + delimiter;
                    str += _estStat[i].dealCount.ToString() + delimiter;
                    str += _estStat[i].symbol + delimiter;
                    str += _estStat[i].tfType + delimiter;
                    str += _estStat[i].tfPeriod.ToString() + delimiter;
                    str += _estStat[i].total.ToString() + delimiter;
                    str += _estStat[i].normTotal.ToString() + delimiter;

                    for ( int j=0; j< _estStat[i].userParam.Count; j++)
                    {
                        str += _estStat[i].userParam[j].value.ToString() + delimiter;
                    }
                    str += _estStat[i].paramTotal.ToString() + delimiter;
                    str += _estStat[i].paramNorm.ToString();
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
            DesParamStratetgy.Version = "2";
            DesParamStratetgy.DateRelease = "16.01.2023";
            DesParamStratetgy.DateChange = "24.01.2023";
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "EstimationStatNorm";
        }
    }
}
