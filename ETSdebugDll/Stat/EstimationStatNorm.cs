using ScriptSolution.Model.Statistics;
using ScriptSolution;
using SourceEts.CommonSettings.Testing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data;
using ScriptSolution.Model.OptimizationResul;
using Export;
using Export.ModelsExport;
using Export.Models;
using Export.Models.Charts;
using System.Windows;
using System.Xml.Linq;
using System.Net;
using Export.Enums;
using Export.DrawingCharts;
using System.Drawing;
using static Export.DrawingCharts.LineETS;
using Point = Export.DrawingCharts.LineETS.Point;
using SourceEts;

namespace EstimationStatNorm
{
    public class EstimationStatNorm : StatisticScript
    {
        StatContainer? _stat = null;
        public ClientReport? _totalPDF = null;
        public List<Action>? _totalActions = null;
        public DateTime _startDate = DateTime.Now; // Начало периода оптимизации
        public DateTime _endDate = DateTime.Now; // Окончание периода оптимизации
        public double _capital = 0; // Начальный капитал
        public double _userSetCommis = 0; // Настройки комиссии на сделку
        public double _userSetSlip = 0; // Настройки проскальзывания на сделку
        public string _userPosSize = ""; // Размер позиции
        public bool _forvard = false;

        /// <summary>
        ///  Начало цикла оптимизации
        /// </summary>
        public override void StartOptimization()
        {
            string path = PathSaveResult;
            path += "\\" + ParamOptimStrategy.NameStrategy + "\\";

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            _totalPDF = new ClientReport();
            _totalPDF.SetExport(new Pdf(path, "Total", true ));
            _totalActions = new List<Action>();
            _stat = new StatContainer(this);
        }
        /// <summary>
        ///  Расчет оценки Форвард интервалов
        /// </summary>
        List<Forvard> CalcForvardEstimate()
        {          
            List<Forvard> result = new List<Forvard>();
            
            foreach (ForvardReportModel frModel in ForvardReports)
            {
                if (frModel.DetailsForvard.Count <= 0)
                    continue;

                Forvard f = new Forvard();
                f._typeIs = frModel.PeriodTypeInSample.ToString();
                f._typeOs = frModel.PeriodTypeOutOfSample.ToString();
                f._periodIs = frModel.PeriodInSample;
                f._periodOs = frModel.PeriodOutOfSample;
                f._estimate = 0;

                foreach( ReportOptimModel mdl in frModel.DetailsForvard )
                {
                    Statistic inS = mdl.StatisticsInSample;
                    Statistic outS = mdl.StatisticsOutOfSample;

                    if (inS.YearProfitLossPercent == 0 || inS.ProfitDealsPercents == 0 ||
                        inS.MaxDrownDownPercent == 0  || inS.FactorRecovery == 0 || inS.CountDeals == 0 )                        
                    {
                        continue;
                    }
                    double kPeriod = f._periodIs / f._periodOs;
                    f._estimate += (outS.YearProfitLossPercent / inS.YearProfitLossPercent +
                    outS.ProfitDealsPercents / inS.ProfitDealsPercents +
                    outS.MaxDrownDownPercent / inS.MaxDrownDownPercent +
                    outS.FactorRecovery / inS.FactorRecovery + 
                    outS.CountDeals * kPeriod / inS.CountDeals );
                }
                f._estimate /= frModel.DetailsForvard.Count;
                result.Add(f);
            }
            return result;
        }
        /// <summary>
        ///  Окончание каждого прохода оптимизации
        /// </summary>
        public override void SetUserStatisticParamOnEndTest( ReportEndOptimForStatistic report )
        {
            if (_stat == null)
                return;

            var oParams = OptimParams;
            var uStat = UserStatistics;
            Statistic stat = report.Statistic;
            _startDate = stat.StartTime;
            _endDate = stat.EndTime;
            _capital = stat.InitialCapital;
            _userSetCommis = report.Comission;
            _userSetSlip = report.Slippage;
            _userPosSize = report.PosSize;

            if ( 
                stat.NetProfitLossPercent >= uStat[0].Value  && 
                Math.Abs( stat.MaxDrownDownPercent ) <= uStat[1].Value &&
                stat.FactorRecovery >= uStat[2].Value  &&
                stat.ProfitDealsPercents >= uStat[3].Value
               )
            {
                EstimationStat est = new EstimationStat();

                if (ForvardReports.Count > 0)
                {
                    est._forvard = CalcForvardEstimate();
                    _forvard = true;
                }
                else
                    _forvard = false;

                est.equity = stat.EquityPoint;
                est.profit = stat.NetProfitLossPercent;
                est.averageDeal = stat.AvaregeProfitLossPercent;
                est.recoveryFactor = stat.FactorRecovery;
                est.dealCount = stat.TradeDealsList.Count;
                est.drawDown = Math.Abs(stat.MaxDrownDownPercent);
                est.drawDownAbs = Math.Abs(stat.MaxAbsDrownDownPercent);
                est.commission = stat.TotalComission;
                est.slippage = stat.TotalSlippage;
                est.yearProfit = stat.YearProfitLossPercent;
                est.profitDeals = stat.ProfitDealsPercents;
                est.symbol = report.Symbol;

                if (report.TimeFrameType == EnumTimeFrame.Minute )
                    est.tfType = "Minute";
                else if (report.TimeFrameType == EnumTimeFrame.Day)
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
                est.CalcInRow( stat );
                _stat.AddStat( est );
            }
        }
        /// <summary>
        ///  Окончание цикла оптимизации
        /// </summary>
        public override List<Dictionary<string, double>> EndOptimizationCycle()
        {
            var uStat = UserStatistics;
            List<Dictionary<string, double>> resToPortfolio = new List<Dictionary<string, double>>();
            EstimationStat result = _stat.BuildReport();

            if ( uStat[4].ValueBool && result != null && result.userParam != null  )
            {
                Dictionary<string, double> parDict = new Dictionary<string, double>();

                foreach ( UserParam p in result.userParam )
                {
                    parDict.Add( p.name, p.value );
                }
                resToPortfolio.Add( parDict );
            }
            _stat = new StatContainer(this);
            return resToPortfolio;
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
            list.Add(new UserParamModel { Name = "Profit ( % ) >= ", Value = 1, Weight = 1,  IsOptimization = false });
            list.Add(new UserParamModel { Name = "Drawdown ( % ) <= ", Value = 50, Weight = 1, IsOptimization = false });
            list.Add(new UserParamModel { Name = "Recovery Factor >= ", Value = 1, Weight = 1, IsOptimization = false });
            list.Add(new UserParamModel { Name = "Average profit ( % ) >= ", Value = 0.3, Weight = 1, IsOptimization = false });
            list.Add(new UserParamModel { Name = "Добавить в портфель", ValueBool = false, IsBool = true, IsOptimization = false });
            return list;
        }

        public override void GetAttributes()
        {
            DesParamStratetgy.Version = "21";
            DesParamStratetgy.DateRelease = new DateTime(2023, 1, 16);
            DesParamStratetgy.DateChange = new DateTime(2023, 6, 13 );
            DesParamStratetgy.Author = "1EX: Vladimir Chernianskiy";
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.LinkFullDescription = "";
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
        public int maxLossCountInRow = 0;
        public int maxProfCountInRow = 0;

        public double profit = 0; // Профит ( % )
        public double drawDown = 0; // Максимальная просадка ( % )
        public double drawDownAbs = 0; // Абсолютная просадка ( % )
        public double recoveryFactor = 0; // Факттор восстановления
        public double averageDeal = 0; // Средняя прибыль на сделку ( % )
        public double dealCount = 0; // Количество сделок 
        public double commission = 0; // комиссия 
        public double slippage = 0;// Проскальзывание 
        public double yearProfit = 0;// Среднегодовая прибыл ( % )
        public double profitDeals = 0;// Процент прибыльных сделок
        public string symbol = ""; // Инструмент
        public string tfType = ""; // Тип ТФ
        public int tfPeriod = 0; // Период ТФ
        public double total = 0; // Общий критерий оценки
        public double normTotal = 0; // Нормированный Общий критерий оценки
        public double paramTotal = 0; // Общий фактор по параметрам
        public double paramNorm = 0; // Нормированный критерий по параметроам
        public List<double>? equity = null; // Набор точек эквити накопительным итогом
        public List<UserParam>? userParam = null;// Параметры оптимизации
        public int targetIdx = -1; // 
        public List<Forvard>? _forvard = null; // Результаты оцненок форвард для данного слоя
 
        public EstimationStat()
        {
            userParam = new List<UserParam>();
            _forvard = new List<Forvard>();
        }
        /// <summary>
        ///  СРасчет рядных показателей
        /// </summary>
        public void CalcInRow(Statistic stat)
        {
            maxLossCountInRow = 0;
            maxProfCountInRow = 0;
            int lossCount = 0;
            int profCount = 0;

            foreach( var d in stat.TradeDealsList )
            {
                if( d.ProfitLoss > 0 )
                {
                    if( lossCount > 0 )
                    {
                        if( lossCount > maxLossCountInRow )
                            maxLossCountInRow = lossCount;
                        lossCount = 0;
                    }
                    profCount++;
                }
                else
                {
                    if (profCount > 0)
                    {
                        if (profCount > maxProfCountInRow)
                            maxProfCountInRow = profCount;
                        profCount = 0;
                    }
                    lossCount++;
                }
            }
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
        bool _addStrategyInfoTotal = true;
        List<EstimationStat> _statContainer;       
        string[] _statNames = { "Profit", "DD", "Recovery", "Avg. Deal", "Deal Count" };

        public StatContainer( EstimationStatNorm mdl )
        {
            _mdl = mdl; 
            _statContainer = new List<EstimationStat>();            
            _totalPDF = _mdl._totalPDF;
            _totalActions = _mdl._totalActions;
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
        ///  Добавляем статистку по проходу
        /// </summary>
        public void AddStat(EstimationStat stat)
        {
            _statContainer.Add(stat);
        }
        /// <summary>
        /// Оценка результата по максимальной интегрированной оценке 
        /// </summary>
        EstimationStat? ResultEstimationMax( List<EstimationStat> statList )
        {
            if (statList.Count <= 0)
                return null;
            
            EstimationStat? result = null;
            double maxEstimate = double.MinValue;
            
            foreach (EstimationStat stat in statList)
            {
                if (stat.normTotal > maxEstimate)
                {
                    maxEstimate = stat.normTotal;
                    result = stat;
                }
            }
            return result;
        }

       /// <summary>
       /// Оценка результата по кластеру распределения 
       /// </summary>
       EstimationStat ResultEstimationByClaster( List<EstimationStat> statList, int rang,  double kLimit )
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
            List<EstimationStat> estStat = _statContainer;

            if (estStat.Count == 0)
                return;

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
        /// <summary>
        /// Рсчет нормализованного стат. фактора 
        /// </summary>
        void CaclNormParams()
        {
            List<EstimationStat> estStat = _statContainer;

            if (estStat.Count == 0)
                return;

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
        /// <summary>
        /// Результирующий отчет
        /// </summary>
        public EstimationStat? BuildReport()
        {
            if (_statContainer.Count == 0)
                return null;

            string path = _mdl.PathSaveResult;
            path += "\\" + _mdl.ParamOptimStrategy.NameStrategy + "\\";
            
            if(!Directory.Exists(path))
                Directory.CreateDirectory(path);

            CaclNormStat();
            CaclNormParams();
            EstimationStat? result = ResultEstimationByClaster(_statContainer, 3, 0.5);
            string fName = _statContainer[0].symbol + "_";
            fName += _statContainer[0].tfType + "_";
            fName += _statContainer[0].tfPeriod;
            ExportReportExcell(_statContainer, path, fName);
            ExportReportPDF( _statContainer, result, path, fName );
            return result;
        }
        /// <summary>
        /// Поиск максимальной оценки форвард 
        /// </summary>
        Forvard ForvardByMaxEvaluation(List<EstimationStat> estStat )
        {
            if ( !_mdl._forvard)
                return null;

            double eval = double.MinValue;
            Forvard? res = null;

            foreach (EstimationStat stat in estStat)
            {
                foreach( Forvard f in stat._forvard )
                {
                    if( f._estimate > eval )
                    {
                        eval = f._estimate;
                        res = f;
                    }
                }
            }
            return  res;
        }
        /// <summary>
        /// Рисуем Equity
        /// </summary>
        void EquityLineChart(EstimationStat result, ClientReport rep, List<Action> actions )
        {
            if (result == null || result.equity == null || result.equity.Count == 0 )
                return; 

            LineChartSet set = new LineChartSet();
            List<LineChartData> data = new List<LineChartData>();
            set.xText = "Период времени";
            set.yText = "Доходность";
            LineChartData equity = new LineChartData();
            equity.width = 2;
            equity.color = Color.Blue;
            DateTime dt = _mdl._startDate;
            TimeSpan delta = _mdl._endDate - dt;
            double dayStep = (double)delta.Days / (double)result.equity.Count;
            double minStep = (double)delta.TotalMinutes / (double)result.equity.Count;
            double minCount = 0;
            double dayCount = 0;

            for (int i = 0; i < result.equity.Count; i++)
            {
                Point p = new Point(i, result.equity[i] - _mdl._capital );
                equity.points.Add(p);

                if (result.tfType == "Day")
                {
                    set.date.Add(dt.AddDays((int)dayCount));
                    dayCount += dayStep;
                }
                else
                {
                    set.date.Add(dt.AddMinutes((int)minCount));
                    minCount += minStep;
                }
            }
            data.Add(equity);
            actions.Add(() => rep.AddChart(new Chart(new Line(new LineETS(set, data)))));
        }
        /// <summary>
        /// Тест линейного графика
        /// </summary>
        void TestLineChart(ClientReport rep, List<Action> actions )
        {
            LineChartSet set = new LineChartSet();
            List<LineChartData> data = new List<LineChartData>();
            set.xText = "Период времени";
            set.yText = "Доходность";
            LineChartData sin1 = new LineChartData();
            LineChartData sin2 = new LineChartData();
            sin1.width = 2;
            sin1.color = Color.Blue;
            sin2.width = 2;
            sin2.color = Color.Green;
            DateTime dt = DateTime.Now;

            int imax = 100; //число точек в периоде
            int t = 2; //число периодов
            int amp1 = 70; //амплитуда
            int amp2 = 30; //амплитуда

            for (int i = 0; i < imax * t; i++)
            {
                Point p = new Point(i, Math.Round(amp1 * Math.Sin(2 * Math.PI / imax * i)));
                sin1.points.Add(p);
                p = new Point(i, Math.Round(amp2 * Math.Sin(2 * Math.PI / imax * i)));
                sin2.points.Add(p);
                set.date.Add(dt.AddDays(i + 1));
            }

            data.Add(sin1);
            data.Add(sin2);
            actions.Add(() => rep.AddChart(new Chart( new Line(new LineETS(set, data)))));
        }
        /// <summary>
        /// Экспорт PDF через Export.dll
        /// </summary>
        void ExportReportPDF( List<EstimationStat> estStat, EstimationStat result, string path, string fName )
        {
            if (estStat.Count <= 0 ) 
                return;

            //Подготовка текстового блока 
            SettingText setTxtJustify = new SettingText();
            setTxtJustify.FontSize = 12;
            setTxtJustify.FontName = "Arial";
            setTxtJustify.TextAligment = Export.Enums.Aligment.Justify;

            SettingText setTxt = new SettingText();
            setTxt.FontSize = 12;
            setTxt.FontName = "Arial";
            setTxt.TextAligment = Export.Enums.Aligment.Left;

            SettingText setTxtCenter = new SettingText();
            setTxtCenter.FontSize = 12;
            setTxtCenter.FontName = "Arial";
            setTxtCenter.TextAligment = Export.Enums.Aligment.Center;

            SettingText setTxtBold = new SettingText();
            setTxtBold.FontSize = 12;
            setTxtBold.FontName = "Arial";
            setTxtBold.TextAligment = Export.Enums.Aligment.Left;
            setTxtBold.Bold = true;
           
            string strategyName = "Наименование стратегии: " + _mdl.ParamOptimStrategy.NameStrategy + "\n";
            string Version = "Версия: " + _mdl.ParamOptimStrategy.Version;
            string Author = "Автор: " + _mdl.ParamOptimStrategy.Author;
            string dateModify = "Дата последней модификации: " + _mdl.ParamOptimStrategy.DateChange;
            string testPeriod = "Период тестирования: " + _mdl._startDate.ToString() + " - " + _mdl._endDate.ToString();
            string startCapital = "Стартовый капитал: " + _mdl._capital;
            string posSize = "Размер позиции: " + _mdl._userPosSize + "\n";
            string symbTF = estStat[0].symbol + " " + estStat[0].tfType + " " + estStat[0].tfPeriod.ToString() + "\n";
            string noResult = "По текущему инcтрументу на данном временном периоде система не имеет устойчивых показателей.";
            ClientReport rep = new ClientReport();
            rep.SetExport(new Pdf(path, fName, false));

            // Если отсутствует результат
            if( result == null )
            {
                rep.GenerateReport(new List<Action>()
                {
                    () => rep.AddText(new Text(strategyName, setTxt)),
                    () => rep.AddText(new Text(Version, setTxt)),
                    () => rep.AddText(new Text(Author, setTxt)),
                    () => rep.AddText(new Text(dateModify, setTxt)),
                    () => rep.AddText(new Text("\n")),
                    () => rep.AddText(new Text(symbTF, setTxtBold)),
                    () => rep.AddText(new Text("\n")),
                    () => rep.AddText(new Text(noResult, setTxt)),
                });
                rep.SaveDocument();

                if (_addStrategyInfo)
                {
                    _totalActions.Add(() => _totalPDF.AddText(new Text(strategyName, setTxt)));
                    _totalActions.Add(() => _totalPDF.AddText(new Text(Version, setTxt )));
                    _totalActions.Add(() => _totalPDF.AddText(new Text(Author, setTxt)));
                    _totalActions.Add(() => _totalPDF.AddText(new Text(dateModify, setTxt)));
                    _totalActions.Add(() => _totalPDF.AddText(new Text("\n")));
                    _addStrategyInfo = false;
                }
                _totalActions.Add(() => _totalPDF.AddText(new Text(symbTF, setTxtBold, true, 1 )));
                _totalActions.Add(() => _totalPDF.AddText(new Text(noResult, setTxt)));
                _totalActions.Add(() => _totalPDF.AddNewPage());
                return;
            }           
            // Сортируем данные и готвим объекты для отчета
            List<EstimationStat> sort = estStat.OrderBy(x => x.paramNorm).ToList();
            string chartName = "Диаграмма плотности распределения оценки результатов оптимизации\n";
            string chartDesc = "Диаграмма показывает зависимость интегрированного показателя статистической оценки " +
                " относительно совокупного фактора параметров оптимизации. Значения по горизонтальной оси - это индекс прохода в результирующей таблице, которая представлена" +
                " в " + fName + ".xlsx" + " файле. Данные в таблице упорядочены по фактору параметров.\n";
            string strResult = "По итогам оптимизации рекомендованы к использованию следующие параметры:\n";
            
            //данные для таблицы выбранных параметров. 
            HeaderTable htbl = new HeaderTable();
            htbl.Headers = new List<string>{ "Наименование", "Значение"};
            TableSetting tableSet = new TableSetting();
            tableSet.BodySetting.SettingText.TextAligment = Export.Enums.Aligment.Center;
            TableModel tableMdl = new TableModel( htbl, tableSet, new List<List<object>>());

            for (int i = 0; i < result.userParam.Count; i++)
            {
                 tableMdl.TableData.Add(new List<object>()
                 {
                     result.userParam[i].name,
                     result.userParam[i].value
                 });
            }
            //данные для таблицы стат. показателей
            HeaderTable hTblStat = new HeaderTable();
            hTblStat.Headers = new List<string> { "", "" };
            TableSetting tableSetStat = new TableSetting();
            tableSetStat.BodySetting.SettingText.TextAligment = Export.Enums.Aligment.Left;
            tableSetStat.TableBorderSetting = new TableBorderSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            tableSetStat.TableBorderInsideSetting = new TableBorderInsideSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            TableModel tableMdlStat = new TableModel(hTblStat, tableSetStat, new List<List<object>>());

            tableMdlStat.TableData.Add(new List<object>()
            {
                "Доходность за период ( % )",
                Math.Round( result.profit, 2 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Среднегодовая доходность ( % )",
                Math.Round(result.yearProfit, 2 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Средняя прибыль на сделку ( % )",
                Math.Round(result.averageDeal, 4),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Процент прибыльных сделок ( % )",
                Math.Round(result.profitDeals, 2 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Количество сделок",
                result.dealCount,
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Максимальная просадка ( % )",
                 Math.Round( result.drawDown, 2 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Абсолютная просадка ( % )",
                 Math.Round( result.drawDownAbs, 2 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "комиссия на сделку ( % )",
                Math.Round( _mdl._userSetCommis, 2 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Проскальзывание на сделку ( пункты )",
                Math.Round( _mdl._userSetSlip, 4 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Общая комиссия ( валюта )",
                Math.Round( result.commission, 4 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Общее проскальзывание ( валюта )",
                Math.Round( result.slippage, 4 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Количество убыточных сделок подряд",
                result.maxLossCountInRow,
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Количество прибыльных сделок подряд",
                result.maxProfCountInRow,
            });

            // Данные по результатам форвард оптимизации
            Forvard f = ForvardByMaxEvaluation(estStat);
            string forvardResult = "";

            if( f != null)
            {
                forvardResult += "In sample (" + f._typeIs + "): " + f._periodIs.ToString() + ";  ";
                forvardResult += "Out of sample (" + f._typeOs +"): " + f._periodOs.ToString() + ";  ";
                forvardResult += "Коэф. устойчивости: " + Math.Round(f._estimate, 2).ToString() + "\n\n";
                forvardResult += "Out of sample - рекомендуемый торговый период, на котором параметры сохраняют свою актуальность.\n";
                forvardResult += "In sample - рекомендуемый минимальный размер обучающей выборки для переоптимизации системы по истечении торгового периода.";
            }
            // Заполняем данные для диаграммы
            List<double> chartData = new List<double>();

            for (int i = 0; i < sort.Count; i++)
            {
                chartData.Add(sort[i].normTotal);
            }
            int markerStart = 0;
            int markerCount = 0;
            // Маркер для выбранного кластера
            if ( result.targetIdx > 0 )
            {
                markerStart = result.targetIdx - 3;
                markerCount = 7; 
            }
            //Настройки диаграммы
            SettingChart setChart = new SettingChart();
            setChart.MarkerSetting.MarkerStart = markerStart;
            setChart.MarkerSetting.MarkerCount = markerCount;
            setChart.SignatureX = "Индекс прохода в текущем цикле";
            setChart.SignatureY = "Интегрированная оценка результата";
            setChart.SettingText.FontSize = 6;
            
            // создание отчета
            List < Action> act = new List<Action>();
            act.Add(() => rep.AddText(new Text(strategyName, setTxtBold)));
            act.Add(() => rep.AddText(new Text(Version, setTxt)));
            act.Add(() => rep.AddText(new Text(Author, setTxt)));
            act.Add(() => rep.AddText(new Text(dateModify, setTxt)));
            act.Add(() => rep.AddText(new Text(testPeriod, setTxt)));
            act.Add(() => rep.AddText(new Text(startCapital, setTxt)));
            act.Add(() => rep.AddText(new Text(posSize, setTxt)));
            act.Add(() => rep.AddText(new Text(symbTF, setTxtBold)));
            act.Add(() => rep.AddText(new Text(chartName, setTxtCenter)));
            act.Add(() => rep.AddChart(new Chart(new Histogram(chartData, setChart))));
            act.Add(() => rep.AddText(new Text(chartDesc, setTxtJustify, false, 1, new HyperLink()
            {
                LinkText = fName + ".xlsx",
                TargetLink = path + "\\" + fName + ".xlsx"
            })));
            act.Add(() => rep.AddText(new Text(strResult, setTxt)));
            act.Add(() => rep.AddTable(tableMdl));
            act.Add(() => rep.AddText(new Text("Показатели статистики, соответствущие рекомендованным параметрам.\n", setTxtBold)));
            act.Add(() => rep.AddTable(tableMdlStat));
            EquityLineChart(result, rep, act);

            if (_mdl._forvard )
            {
                act.Add(() => rep.AddText(new Text("По результатам форвард-оптимизации рекомендуются следующие интервалы:\n", setTxtBold)));
                act.Add(() => rep.AddText(new Text(forvardResult, setTxt)));
            }
            rep.GenerateReport(act);
            rep.SaveDocument();
             
            //Добавляем данные в глобальный отчет
            if( _addStrategyInfoTotal)
            {
                _totalActions.Add(() => _totalPDF.AddText(new Text(strategyName, setTxt)));
                _totalActions.Add(() => _totalPDF.AddText(new Text(Version, setTxt)));
                _totalActions.Add(() => _totalPDF.AddText(new Text(Author, setTxt)));
                _totalActions.Add(() => _totalPDF.AddText(new Text(dateModify, setTxt)));
                _totalActions.Add(() => _totalPDF.AddText(new Text(testPeriod, setTxt)));
                _totalActions.Add(() => _totalPDF.AddText(new Text(startCapital, setTxt)));
                _totalActions.Add(() => _totalPDF.AddText(new Text(posSize, setTxt)));
                _addStrategyInfoTotal = false;
            }
            _totalActions.Add(() => _totalPDF.AddText(new Text(symbTF, setTxtBold, true, 1)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(chartName, setTxtCenter)));
            _totalActions.Add(() => _totalPDF.AddChart(new Chart(new Histogram(chartData, setChart))));
            //_totalActions.Add(() => _totalPDF.AddChart(new Chart(new Histogram(chartData, markerStart, markerCount))));
            _totalActions.Add(() => _totalPDF.AddText(new Text(chartDesc, setTxtJustify, false, 1, new HyperLink()
            {
                LinkText = fName + ".xlsx",
                TargetLink = path + "\\" + fName + ".xlsx"
            })));
            _totalActions.Add(() => _totalPDF.AddText(new Text(strResult, setTxt)));
            _totalActions.Add(() => _totalPDF.AddTable(tableMdl));
            _totalActions.Add(() => _totalPDF.AddText(new Text("Показатели статистики, соответствущие рекомендованным параметрам.\n", setTxtBold)));
            _totalActions.Add(() => _totalPDF.AddTable(tableMdlStat));
            EquityLineChart(result, _totalPDF, _totalActions);

            if (_mdl._forvard)
            {
                _totalActions.Add(() => _totalPDF.AddText(new Text("По результатам форвард-оптимизации рекомендуются следующие интервалы:\n", setTxtBold)));
                _totalActions.Add(() => _totalPDF.AddText(new Text(forvardResult, setTxt)));
            }
            _totalActions.Add(() => _totalPDF.AddNewPage());
        }
        /// <summary>
        /// Экспорт Excell 
        /// </summary>
        void ExportReportExcell(List<EstimationStat> estStat, string path, string fName)
        {
            if (estStat.Count <= 0)
                return;
            // Данные по оптимизации
            HeaderTable htbl = new HeaderTable();
            htbl.Headers = new List<string>();
            htbl.Headers.Add("Symbol");
            htbl.Headers.Add("Time Type");
            htbl.Headers.Add("Time");

            for (int i = 0; i < _statNames.Length; i++)
            {
                htbl.Headers.Add(_statNames[i]);
            }

            htbl.Headers.Add("Stat Norm");

            for (int i = 0; i < estStat[0].userParam.Count; i++)
            {
                htbl.Headers.Add(estStat[0].userParam[i].name);
            }
            htbl.Headers.Add("Param. Norm");
            // Сортируем данные и готвим объекты для отчета
            List<EstimationStat> sort = estStat.OrderBy(x => x.paramNorm).ToList();
            TableSetting tableSetStat = new TableSetting();
            tableSetStat.BodySetting.SettingText.TextAligment = Export.Enums.Aligment.Center;
            tableSetStat.TableBorderSetting = new TableBorderSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            tableSetStat.TableBorderInsideSetting = new TableBorderInsideSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            TableModel tableMdl = new TableModel(htbl, tableSetStat, new List<List<object>>());

            for (int i = 0; i < sort.Count; i++)
            {
                List<object> tObjects = new List<object>();
                tObjects.Add(sort[i].symbol);
                tObjects.Add(sort[i].tfType);
                tObjects.Add(sort[i].tfPeriod);
                tObjects.Add(sort[i].profit);
                tObjects.Add(sort[i].drawDown);
                tObjects.Add(sort[i].recoveryFactor);
                tObjects.Add(sort[i].averageDeal);
                tObjects.Add(sort[i].dealCount);
                tObjects.Add(sort[i].normTotal);

                for (int j = 0; j < sort[i].userParam.Count; j++)
                {
                    tObjects.Add(sort[i].userParam[j].value);
                }
                tObjects.Add(sort[i].paramNorm.ToString());
                tableMdl.TableData.Add(tObjects);
            }
            
            // Данные по форвард
            HeaderTable htblForvard = new HeaderTable();
            htblForvard.Headers = new List<string>();
            TableModel? tableMdlForvard = null;

            if (_mdl._forvard )
            {
                for (int i = 0; i < estStat[0].userParam.Count; i++)
                {
                    htblForvard.Headers.Add(estStat[0].userParam[i].name);
                }
                htblForvard.Headers.Add("PeriodIn(" + sort[0]._forvard[0]._typeIs + ")");
                htblForvard.Headers.Add("PeriodOut(" + sort[0]._forvard[0]._typeOs + ")");
                htblForvard.Headers.Add("Evaluation");
                tableMdlForvard = new TableModel(htblForvard, tableSetStat, new List<List<object>>());

                for (int i = 0; i < sort.Count; i++)
                {
                    foreach (Forvard f in sort[i]._forvard)
                    {
                        List<object> tObjectsForvard = new List<object>();

                        for (int j = 0; j < sort[i].userParam.Count; j++)
                        {
                            tObjectsForvard.Add(sort[i].userParam[j].value);
                        }
                        tObjectsForvard.Add(f._periodIs);
                        tObjectsForvard.Add(f._periodOs);
                        tObjectsForvard.Add(f._estimate);
                        tableMdlForvard.TableData.Add(tObjectsForvard);
                    }
                }
            }
            // Выод отчета
            ClientReport rep = new ClientReport();
            rep.SetExport(new Excel(path, fName, "Optimization"));
            List<Action> act = new List<Action>();
            act.Add(() => rep.AddTable(tableMdl));

            if (_mdl._forvard )
            {
                act.Add(() => rep.AddNewPage("Forvard"));
                act.Add(() => rep.AddTable(tableMdlForvard));
            }
            rep.GenerateReport(act);
            rep.SaveDocument();
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
                    str = "По итогам оптимизации рекомендованы к использованию следующие параметры:";
                    sw.WriteLine(str);

                    for (int i = 0; i < result.userParam.Count; i++)
                    {
                        str = result.userParam[i].name + " = " + result.userParam[i].value.ToString();
                        sw.WriteLine(str);
                    }
                }
                else
                {
                    str = "Система неустойчива на всем диапазоне оптимизируемых параметров.";
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
    /// <summary>
    /// FRVARD ОПТИМИЗАЦИЯ
    /// </summary>
    public class Forvard
    {
        public string _typeIs = "";
        public int _periodIs = 0;
        public string _typeOs = "";
        public int _periodOs = 0;
        public double _estimate = 0;
    }
}
