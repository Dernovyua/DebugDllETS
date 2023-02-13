using ScriptSolution.Model.Statistics;
using ScriptSolution;
using SourceEts.CommonSettings.Testing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data;
//using System.Windows.Media;
using ScriptSolution.Model.OptimizationResul;
using Export;
using Export.ModelsExport;
using Export.Models;
using Export.Models.Charts;
using System.Windows;

using System.Drawing.Imaging;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace EstimationStatNorm
{
    public class EstimationStatNorm : StatisticScript
    {
                StatContainer? _stat = null;
        //public  ClientReport _clientRep = default!;

        /// <summary>
        ///  Начало цикла оптимизации
        /// </summary>
        public override void StartOptimization()
        {
            _stat = new StatContainer( this );
            //_clientRep = new ClientReport();
        }
        /// <summary>
        ///  Окончание каждого прохода оптимизации
        /// </summary>
        public override void SetUserStatisticParamOnEndTest( ReportEndOptimForStatistic report )
        {
            var oParams = OptimParams;
            var uStat = UserStatistics;
            Statistic stat = report.Statistic; 
            
            if( 
                stat.NetProfitLossPercent >= uStat[0].Value  && 
                Math.Abs( stat.MaxDrownDownPercent ) <= uStat[1].Value &&
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
        public override void EndOptimizationCycle()
        {
            _stat.BuildReport();            
        }
        /// <summary>
        ///  Окончание процеса оптимизации
        /// </summary>
        public override void EndOptimizationAll()
        {
            _stat.OnEndOptimisation();
        }
        /// <summary>
        /// Определение параметров для скрипта
        /// </summary>
        public override List<UserParamModel> CreateUserStatisticParam()
        {
            List<UserParamModel> list = new List<UserParamModel>();
            list.Add(new UserParamModel { Name = "Profit ( % ) >= ", Value = 1, IsOptimization = false });
            list.Add(new UserParamModel { Name = "Drawdown ( % ) <= ", Value = 50, IsOptimization = false });
            list.Add(new UserParamModel { Name = "Recovery Factor >= ", Value = 1, IsOptimization = false });
            list.Add(new UserParamModel { Name = "Deal profit ( % ) >= ", Value = 0.3, IsOptimization = false });
            return list;
        }

        public override void GetAttributes()
        {
            DesParamStratetgy.Version = "7";
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
        public int targetIdx = -1;

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
        EstimationStatNorm? _mdl = null;
        ClientReport? _totalPDF = null;
        List<Action>? _totalActions = null;
        bool _addStrategyInfo = true;
        List<List<EstimationStat>> _statContainer;       
        string[] _statNames = { "Profit", "DD", "Recovery", "Avg. Deal", "Deal Count" };
        int _reportCount = 0;

        public StatContainer( EstimationStatNorm mdl )
        {
            _mdl = mdl; 
            _statContainer = new List<List<EstimationStat>>();
            _totalPDF = new ClientReport();
            _totalPDF.SetExport(new Pdf(_mdl.PathSaveResult, "total"));
            _totalActions = new List<Action>();
        }
        /// <summary>
        /// Окончание процесса оптимизации
        /// </summary>
        public void OnEndOptimisation()
        {
            _totalPDF.GenerateReport( _totalActions );
            _totalPDF.SaveDocument();
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
        /// Оценка результата
        /// </summary>
        EstimationStat ResultEstimation( List<EstimationStat> statList, int rang,  double kLimit )
        {
            if (statList.Count < (rang * 2 + 1))
                return null;

            List<EstimationStat> sorted = new List<EstimationStat>();
            sorted = statList.OrderBy(x => x.paramNorm).ToList();
            double maxResult = double.MinValue;
            int targetIdx = -1;

            for( int i=rang+1; i < sorted.Count - ( rang + 1 ); i++ )
            {
                bool targetStat = true;

                for(int j=0; j<rang; j++)
                {
                    double limit = sorted[i].normTotal * kLimit;

                    if ( sorted[i + j].normTotal > sorted[i].normTotal ||
                         sorted[i + j].normTotal < limit ||
                         sorted[i - j].normTotal > sorted[i].normTotal ||
                         sorted[i - j].normTotal < limit
                        )
                    {
                        targetStat = false;
                        break;
                    }
                }

                if (!targetStat)
                    continue;

                if (sorted[i].normTotal > maxResult)
                {
                    maxResult = sorted[i].normTotal;
                    targetIdx = i;
                }
            }

            if (targetIdx >= 0)
            {
                sorted[targetIdx].targetIdx = targetIdx;
                return sorted[targetIdx];
            }
            return null;
        }

        /// <summary>
        /// Нормализация параметра
        /// </summary>
        double CalcNorm(double value, double min, double max)
        {
            double norm = max - min;

            if (norm <= 0)
                return 0; //double.NaN;

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
                    //--- debug
                    if(estStat[i].userParam.Count != 2 )
                    {
                        int dbg = 0;
                    }
                    //---
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
        /// Результирующий отчет
        /// </summary>
        public void BuildReport()
        {
            string path = _mdl.PathSaveResult;            
            CaclNormStat();
            CaclNormParams();

            for (int i = 0; i < _statContainer.Count; i++)
            {
                if (i < _reportCount)
                    continue;

                EstimationStat result = ResultEstimation(_statContainer[i], 4, 0.5);
                string fName = _statContainer[i][0].symbol + "_";
                fName += _statContainer[i][0].tfType + "_";
                fName += _statContainer[i][0].tfPeriod;
                SaveResultCSV(_statContainer[i], result, path + "\\" + fName + ".csv", ";");
                ExportReportPDF( _statContainer[i], result, path, fName );
                //TestPDF(path, fName);
                _reportCount++;
            }
        }
        /// <summary>
        /// Тест экспорта PDF через Export.dll
        /// </summary>
        public void TestPDF(string path, string fName)
        {
            using (ClientReport rep = new ClientReport())
            {
                rep.SetExport(new Pdf(path, fName));
                rep.GenerateReport(new List<Action>()
                {
                    () => rep.AddText(new Text("Helllooooooooooooooooooo", new SettingText())),
                    () => rep.AddChart(new Chart(new Histogram(new List<double> () { 0.1, 0.14, 0.28, 0.9, 0.14, 0.6, 0.44, 0.63 }, new SettingChart()))),
                });

                rep.SaveDocument();
            }
        }
        /// <summary>
        /// Экспорт PDF через Export.dll
        /// </summary>
        void ExportReportPDF( List<EstimationStat> estStat, EstimationStat result, string path, string fName )
        {
            if (estStat.Count <= 0 || result == null )
                return;

            ClientReport rep = new ClientReport();
            rep.SetExport(new Pdf(path, fName));
            List<EstimationStat> sort = estStat.OrderBy(x => x.paramNorm).ToList();

            //Подготовка текстового блока 
            SettingText setTxt = new SettingText();
            setTxt.FontSize = 12;
            setTxt.FontName = "Arial";
            setTxt.TextAligment = Export.Enums.Aligment.Left;

            SettingText setTxtCenter = new SettingText();
            setTxtCenter.FontSize = 12;
            setTxtCenter.FontName = "Arial";
            setTxtCenter.TextAligment = Export.Enums.Aligment.Center;

            SettingText setTxtBold = new SettingText();
            setTxtBold.FontSize = 14;
            setTxtBold.FontName = "Arial";
            setTxtBold.TextAligment = Export.Enums.Aligment.Left;
            setTxtBold.Bold = true;

            string strategyName = "Наименование стратегии: " + _mdl.ParamOptimStrategy.NameStrategy;
            string Version = "Версия: " + _mdl.ParamOptimStrategy.Version;
            string Author = "Автор: " + _mdl.ParamOptimStrategy.Author;
            string dateModify = "Дата последней модификации: " + _mdl.ParamOptimStrategy.DateChange;
            string symbTF = estStat[0].symbol + "\t" + estStat[0].tfType + " " + estStat[0].tfPeriod.ToString();
            string chartName = "Диаграмма плотности распределения оценки результатов оптимизации";
            string chartDesc = "Данная диаграмма показывает зависимость интегрированного показателя статистической оценки " +
                " относительно совокупного фактора парамеров оптимизации. Значения по горизонтальной оси, это индекс прохода в результирующей таблице которая представлена" +
                " в соответствующем CSV файле. Данные в таблице упорядочены по фактору параметров.";
            string strResult = "По итогам оптимизации, рекомендованы к использованию следующие параметры:\n\n";

            //данные для таблицы выбранных параметров. 
            HeaderTable htbl = new HeaderTable();
            htbl.Headers = new List<string>{ "Нименование", "Значение"};
            TableModel tableMdl = new TableModel( htbl, new TableSetting(), new List<List<object>>());

            for (int i = 0; i < result.userParam.Count; i++)
            {
                 tableMdl.TableData.Add(new List<object>()
                 {
                     result.userParam[i].name,
                     result.userParam[i].value
                 });
            }

            // Заполняем данные для диаграммы
            List<double> chartData = new List<double>();

            for (int i = 0; i < sort.Count; i++)
            {
                chartData.Add(sort[i].normTotal);
            }
            /*SettingChart chartSet = new SettingChart();
            chartSet.Width = 800;
            chartSet.Height = 350;
            chartSet.SignatureX = "Индекс прохода";
            chartSet.SignatureY = "Интегрированная оценка";*/
            int markerStart = 0;
            int markerCount = 0;
            // Маркер для выбранного кластера
            if ( result.targetIdx > 0 )
            {
                markerStart = result.targetIdx - 3;
                markerCount = 7; 
            }
            // создание отчета
            rep.GenerateReport(new List<Action>()
            {
                () => rep.AddText(new Text(strategyName, setTxt)),
                () => rep.AddText(new Text(Version, setTxt)),
                () => rep.AddText(new Text(Author, setTxt)),
                () => rep.AddText(new Text(dateModify, setTxt)),
                () => rep.AddText(new Text("\n")),
                () => rep.AddText(new Text(symbTF, setTxtBold)),
                () => rep.AddText(new Text("\n")),
                () => rep.AddText(new Text(chartName, setTxtCenter)),
                //() => rep.AddChart(new Chart(new Histogram( chartData, chartSet ))),
                () => rep.AddChart(new Chart(new Histogram( chartData, markerStart, markerCount ))),
                () => rep.AddText(new Text(chartDesc, setTxt)),
                () => rep.AddText(new Text("\n")),
                () => rep.AddText(new Text(strResult, setTxt)),
                () => rep.AddTable(tableMdl)
             });
             //rep.OpenPreview();
             rep.SaveDocument();
             //Добавляем данные в глобальный отчет
             if( _addStrategyInfo)
             {
                 _totalActions.Add(() => _totalPDF.AddText(new Text(strategyName, setTxt)));
                 _totalActions.Add(() => _totalPDF.AddText(new Text(Version, setTxt)));
                 _totalActions.Add(() => _totalPDF.AddText(new Text(Author, setTxt)));
                 _totalActions.Add(() => _totalPDF.AddText(new Text(dateModify, setTxt)));
                 _totalActions.Add(() => _totalPDF.AddText(new Text("\n")));
                 _addStrategyInfo = false;
             }
             _totalActions.Add(() => _totalPDF.AddText(new Text(symbTF, setTxtBold)));
             _totalActions.Add(() => _totalPDF.AddText(new Text("\n")));
             _totalActions.Add(() => _totalPDF.AddText(new Text(chartName, setTxtCenter)));
             _totalActions.Add(() => rep.AddChart(new Chart(new Histogram(chartData, markerStart, markerCount))));
             _totalActions.Add(() => _totalPDF.AddText(new Text("\n")));
             _totalActions.Add(() => _totalPDF.AddText(new Text(strResult, setTxt)));
             _totalActions.Add(() => _totalPDF.AddTable(tableMdl));
             //_totalPDF.AddNewPage();
        }
        /// <summary>
        /// Сохранение результатов в CSV
        /// </summary>
        void SaveResultCSV( List<EstimationStat> estStat, EstimationStat result, string fNameFull, string delimiter)
        {
            if ( estStat.Count <= 0)
                return;

            using (StreamWriter sw = new StreamWriter(fNameFull, false, System.Text.Encoding.UTF8))
            {
                string str = "";

                //-- Выводим результата анализа 
                sw.WriteLine(" ");

                if (result != null)
                {
                    str = "По итогам оптимизации, рекомендованы к использованию следующие параметры:";
                    sw.WriteLine(str);

                    for (int i = 0; i < result.userParam.Count; i++)
                    {
                        str = result.userParam[i].name + " = " + result.userParam[i].value.ToString();
                        sw.WriteLine(str);
                    }
                }
                else
                {
                    str = "Система не устойчива на всем диапазоне оптимизируемых параметров.";
                    sw.WriteLine(str);
                }
                sw.WriteLine(" ");
                sw.WriteLine(" ");
                str = "";
                //--- Выводим табличный отчет 
                for (int i = 0; i < _statNames.Length; i++)
                {
                    str += _statNames[i] + delimiter;
                }

                str += "Symbol" + delimiter;
                str += "Time Type" + delimiter;
                str += "Time" + delimiter;
                str += "Stat Norm" + delimiter;

                for (int i = 0; i < estStat[0].userParam.Count; i++)
                {
                    str += estStat[0].userParam[i].name + delimiter;
                }
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
                    str += estStat[i].normTotal.ToString() + delimiter;

                    for (int j = 0; j < estStat[i].userParam.Count; j++)
                    {
                        str += estStat[i].userParam[j].value.ToString() + delimiter;
                    }

                    str += estStat[i].paramNorm.ToString();
                    sw.WriteLine(str);
                }
                sw.Close();
            }
        }
    }
}
