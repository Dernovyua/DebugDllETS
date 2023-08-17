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
using SourceEts;
using System.Globalization;

namespace MODEL_EVALUATION
{
    public class ModelEvaluation : StatisticScript
    {
        StatContainer? _stat = null;
        public DateTime _startDate = DateTime.Now; // Начало периода оптимизации
        public DateTime _endDate = DateTime.Now; // Окончание периода оптимизации
        public double _capital = 0; // Начальный капитал
        public double _userSetCommis = 0; // Настройки комиссии на сделку
        public double _userSetSlip = 0; // Настройки проскальзывания на сделку
        public string _userPosSize = ""; // Размер позиции

        /// <summary>
        ///  Начало цикла оптимизации
        /// </summary>
        public override void StartOptimization()
        {
            string path = PathSaveResult;
            path += "\\" + ParamOptimStrategy.NameStrategy + "\\";

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            _stat = new StatContainer(this);
        }
        /// <summary>
        ///  Окончание каждого прохода оптимизации
        /// </summary>
        public override void SetUserStatisticParamOnEndTest(ReportEndOptimForStatistic report)
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

            if ( stat.FactorRecovery >= uStat[0].Value && stat.CountDeals >= uStat[1].Value && stat.ProfitDealsPercents >= uStat[2].Value )
            {
                EstimationStat est = new EstimationStat();
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

                if (report.TimeFrameType == EnumTimeFrame.Minute)
                    est.tfType = "Minute";
                else if (report.TimeFrameType == EnumTimeFrame.Day)
                    est.tfType = "Day";
                else
                    est.tfType = "Unknown";

                est.tfPeriod = report.TimeFramePeriod;

                for (int i = 0; i < oParams.Count; i++)
                {
                    UserParam up = new UserParam();
                    up.name = oParams[i].Name;
                    up.value = oParams[i].Value;
                    est.userParam.Add(up);
                }
                _stat.AddStat(est);
            }
        }
        /// <summary>
        ///  Окончание цикла оптимизации
        /// </summary>
        public override List<Dictionary<string, double>> EndOptimizationCycle()
        {
            var uStat = UserStatistics;
            List<Dictionary<string, double>> resToPortfolio = new List<Dictionary<string, double>>();
            _stat.LoadModelCollection(RobotParams.ParamOptimizations[0].ValueString);
            _stat.BuildReport(RobotParams.ParamOptimizations[0].ValueString );
            _stat = new StatContainer(this);
            return resToPortfolio;
        }
        /// <summary>
        ///  Окончание процеса оптимизации
        /// </summary>
        public override void EndOptimizationAll()
        {
        }
        /// <summary>
        /// Определение параметров для скрипта
        /// </summary>
        public override List<UserParamModel> CreateUserStatisticParam()
        {
            List<UserParamModel> list = new List<UserParamModel>();
            list.Add(new UserParamModel { Name = "Recovery Factor >= ", Value = 2, Weight = 1, IsOptimization = false });
            list.Add(new UserParamModel { Name = "Deal count >= ", Value = 9, Weight = 1, IsOptimization = false });
            list.Add(new UserParamModel { Name = "Frofi deals percent >= ", Value = 60, Weight = 1, IsOptimization = false });
            return list;
        }

        public override void GetAttributes()
        {
            DesParamStratetgy.Version = "1";
            DesParamStratetgy.DateRelease = new DateTime(2023, 1, 16);
            DesParamStratetgy.DateChange = new DateTime(2023, 6, 13);
            DesParamStratetgy.Author = "1EX: Vladimir Chernianskiy";
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.LinkFullDescription = "";
            DesParamStratetgy.NameStrategy = "ModelEvaluation";
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
        public List<UserParam>? userParam = null;// Параметры оптимизации
        public int targetIdx = -1; // 

        public EstimationStat()
        {
            userParam = new List<UserParam>();
        }

    };
    /// <summary>
    /// Модель
    /// </summary>
    public class Model
    {
        public int tCount = 0;
        public double lProb = 0;
        public double sProb = 0;
        public double nProb = 0;
    }
    /// <summary>
    ///  СТАТИСТИКА
    /// </summary>
    public class StatContainer
    {
        ModelEvaluation? _mdl = null;
        List<EstimationStat> _statContainer;
        List<Model> _modelCollection = new List<Model>();
        string[] _statNames = { "Profit", "DD", "Recovery", "Avg. Deal", "Deal Count" };

        public StatContainer(ModelEvaluation mdl)
        {
            _mdl = mdl;
            _statContainer = new List<EstimationStat>();
        }
        /// <summary>
        ///  Загружаем коллекцию моделей
        /// </summary>
        public void LoadModelCollection(string fName)
        {
            _modelCollection = new List<Model>();

            using (StreamReader sr = new StreamReader(fName, System.Text.Encoding.Default))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {
                    string[] words = line.Split('|');
                    
                    if (words[0].Equals("PATTERN"))
                    {
                        Model m  = new Model();
                        m.lProb = Convert.ToDouble(words[1].Replace(',', '.'), CultureInfo.InvariantCulture.NumberFormat);
                        m.sProb = Convert.ToDouble(words[2].Replace(',', '.'), CultureInfo.InvariantCulture.NumberFormat);
                        m.nProb = Convert.ToDouble(words[3].Replace(',', '.'), CultureInfo.InvariantCulture.NumberFormat);
                        m.tCount = int.Parse(words[4]);
                        _modelCollection.Add(m);    
                    }
                }
            }
        }
        /// <summary>
        ///  Добавляем статистку по проходу
        /// </summary>
        public void AddStat(EstimationStat stat)
        {
            _statContainer.Add(stat);
        } 
        /// <summary>
        /// Результирующий отчет
        /// </summary>
        public void BuildReport( string modelPath )
        {
            if (_statContainer.Count == 0)
                return;

            string path = _mdl.PathSaveResult;
            path += "\\" + _mdl.ParamOptimStrategy.NameStrategy + "\\";

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            string fName = _statContainer[0].symbol + "_";
            fName += _statContainer[0].tfType + "_";
            fName += _statContainer[0].tfPeriod;
            ExportReportExcell(_statContainer, path, fName);
            ModifyModelFile(modelPath);
        }
        /// <summary>
        /// Запись результатов оценки в файл модели
        /// </summary>
        void ModifyModelFile( string modelPath )
        {
            string res = "EVALUATION";
            var uStat = _mdl.UserStatistics;

            for ( int i=0; i<_statContainer.Count; i++ )
            {
                int pid = (int)_statContainer[i].userParam[0].value;
                double prc = uStat[2].Value / 100;
                int count = (int)uStat[1].Value;

                if ( _modelCollection[pid].tCount >= count &&
                   ( _modelCollection[pid].lProb >= prc || _modelCollection[pid].sProb >= prc ))
                {
                    res += "|";
                    res += _statContainer[i].userParam[0].value.ToString();
                }
            }
            File.AppendAllText( modelPath, res );
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

            for (int i = 0; i < estStat[0].userParam.Count; i++)
            {
                htbl.Headers.Add(estStat[0].userParam[i].name);
            }
            // Сортируем данные и готвим объекты для отчета
            //List<EstimationStat> sort = estStat.OrderBy(x => x.recoveryFactor).ToList();
            List<EstimationStat> sort = estStat;
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

                for (int j = 0; j < sort[i].userParam.Count; j++)
                {
                    tObjects.Add(sort[i].userParam[j].value);
                }
                tableMdl.TableData.Add(tObjects);
            }

            // Выод отчета
            ClientReport rep = new ClientReport();
            rep.SetExport(new Excel(path, fName, "Optimization"));
            List<Action> act = new List<Action>();
            act.Add(() => rep.AddTable(tableMdl));
            rep.GenerateReport(act);
            rep.SaveDocument();
        }
    }
}
