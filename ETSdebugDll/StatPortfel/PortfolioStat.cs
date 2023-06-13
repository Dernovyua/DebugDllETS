using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Interop;
using Export;
using Export.Enums;
using Export.Models;
using Export.ModelsExport;
using Export.Models.Charts;
using Export.DrawingCharts;
using ScriptSolution;
using ScriptSolution.Model;
using ScriptSolution.Model.Portfels;
using ScriptSolution.Model.Portfels.PortfelTest;
using ScriptSolution.Model.Statistics;
using ScriptSolution.ScanerModel;
using SourceEts.CommonSettings.Testing;
using System.Drawing;
using static Export.DrawingCharts.LineETS;
using Point = Export.DrawingCharts.LineETS.Point;

namespace ETSdebugDll.StatPortfel
{
    public class PortfolioStat : PortfelScript
    {
        public ClientReport _totalPDF = new ClientReport();
        public List<Action> _totalActions = new List<Action>();
        List<int> _exeption = new List<int>(); // Номера роботов подлежащие удалению из портфеля
        List<int> _noDeals = new List<int>(); // Номера роботов не имеющих ни одной сделки на всем периоде
        double _entrySize = 10000;
        double _coreLimit = 0.5;

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
        /// <summary>
        ///  Событие по окончании теста
        /// </summary>
        public override List<List<int>> ExecuteEndOptimization(List<PortfelResultTest> results, List<CorelationModel> corelations)
        {
            // var userParams = UserParams; //пользовательски параметры для скрипта используются только для тестирования
            string path = PathSaveResult;
            string name = "\\PortfolioTest.pdf";
            _totalPDF = new ClientReport();
            _totalPDF.SetExport(new Pdf(path, name, false));
            _totalActions = new List<Action>();
            _exeption = new List<int>();
            _noDeals = new List<int>();

            foreach (PortfelResultTest res in results)
            {
                if( !CheckForEquityPoints(res) )
                {
                    _noDeals.Add(res.NumberRobot);
                    continue;
                }
                AddToReport(res, corelations);
            }
            SettingText setTxtBold = new SettingText();
            setTxtBold.FontSize = 12;
            setTxtBold.FontName = "Arial";
            setTxtBold.TextAligment = Export.Enums.Aligment.Left;
            setTxtBold.Bold = true;
            
            if( _noDeals.Count > 0)
            {
                string noDealsStr = "Элементы со следующими номерами ( ";

                for(int i=0; i<_noDeals.Count; i++)
                {
                    noDealsStr += _noDeals[i].ToString();

                    if (i < _noDeals.Count - 1)
                        noDealsStr += ", ";
                }
                noDealsStr += " ) не имеют сделок на всем периоде теста.\n";
                _totalActions.Add(() => _totalPDF.AddText(new Text( noDealsStr, setTxtBold)));
            }

            if (_exeption.Count > 0)
            {
                _totalActions.Add(() => _totalPDF.AddText(new Text("Рекомендуется удалить из портфеля элементы со следующими номерами.\n", setTxtBold)));
                _totalActions.Add(() => _totalPDF.AddTable(ExeptionTable(corelations)));
            }
            _totalActions.Add(() => _totalPDF.AddText(new Text("Распределение капитала в соответствии с весами.\n", setTxtBold)));
            _totalActions.Add(() => _totalPDF.AddTable(WeightTable(results)));
            _totalPDF.GenerateReport(_totalActions);
            _totalPDF.SaveDocument();
            return new List<List<int>>();
        }
        /// <summary>
        /// Считаем среднюю корреляцию по портфелю
        /// </summary>
        double CalcAveragePortfolioCorrelation(int robotNum, List<CorelationModel> corelations)
        {
            if(corelations.Count <= 1)
                return double.NaN;
            double res = 0;

            foreach (CorelationModel cor in corelations)
            {
                if (robotNum == cor.NumberRobot)
                {                  
                    for(int i=0; i<cor.CorValues.Count; i++)
                    {
                        if (cor.CorValues[i].NumberRobot != robotNum)
                            res += cor.CorValues[i].Corelation;
                    }
                    res /= (cor.CorValues.Count - 1);
                    return res;
                }
            }
            return double.NaN;
        }
        /// <summary>
        /// Проверяем массив эквити на отсутствие сделок 
        /// </summary>
        bool CheckForEquityPoints(PortfelResultTest res)
        {
            bool deals = false;

            for (int i = 0; i < res.Statistic.EquityPoint.Count; i++)
            {
                if (res.Statistic.EquityPoint[i] != 0 && res.Statistic.EquityPoint[i] != double.NaN)
                {
                    deals = true;
                    break;
                }
            }
            return deals;
        }
        /// <summary>
        /// Рисуем Equity
        /// </summary>
        void EquityLineChart(PortfelResultTest res, ClientReport rep, List<Action> actions)
        {
            if (res == null || res.Statistic.EquityPoint == null || res.Statistic.EquityPoint.Count == 0)
                return;

            Statistic stat = res.Statistic;
            LineChartSet set = new LineChartSet();
            List<LineChartData> data = new List<LineChartData>();
            set.xText = "Период времени";
            set.yText = "Доходность";
            LineChartData equity = new LineChartData();
            equity.width = 2;
            equity.color = Color.Blue;
            DateTime dt = stat.StartTime;

            for (int i = 0; i < stat.EquityPoint.Count; i++)
            {
                Point p = new Point( i, stat.EquityPoint[i] - stat.InitialCapital );
                equity.points.Add(p);

                if ( res.TimeFrameType ==  SourceEts.EnumTimeFrame.Day )
                    set.date.Add(dt.AddDays(i));
                else
                    set.date.Add(dt.AddMinutes(i * res.TimeFramePeriod ));
            }
            data.Add(equity);
            actions.Add(() => rep.AddChart(new Chart(new Line(new LineETS(set, data)))));
        }
        /// <summary>
        /// данные для таблицы стат. показателей
        /// </summary>
        TableModel StatisticTable(PortfelResultTest res)
        {
            Statistic stat = res.Statistic;
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
                stat.RealNetProfitLossPercent
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Среднегодовая доходность ( % )",
                stat.YearProfitLossPercent
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Средняя прибыль на сделку ( % )",
                stat.AvaregeProfitLossPercent,
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Процент прибыльных сделок ( % )",
                stat.ProfitDealsPercents
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Количество сделок",
                stat.CountDeals,
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Максимальная просадка ( % )",
                stat.MaxDrownDownPercent
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Максимальная абсолютная просадка ( % )",
                stat.MaxAbsDrownDownPercent
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "комиссия на сделку ( % )",
                Math.Round( res.Comission, 4 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Проскальзывание на сделку ( пункты )",
                Math.Round( res.Slippage, 4 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Общая комиссия ( валюта )",
                Math.Round( stat.TotalComission, 4 ),
            });
            tableMdlStat.TableData.Add(new List<object>()
            {
                "Общее проскальзывание ( валюта )",
                Math.Round( stat.TotalSlippage, 4 ),
            });
            return tableMdlStat;
        }
        /// <summary>
        ///  Модель таблицы исключений
        /// </summary>
        TableModel ExeptionTable(List<CorelationModel> corelations)
        {
            HeaderTable hTblStat = new HeaderTable();
            hTblStat.Headers = new List<string> { " № Робота", "Коэф. корреляции по портфелю" };
            TableSetting tableSetStat = new TableSetting();
            tableSetStat.BodySetting.SettingText.TextAligment = Export.Enums.Aligment.Left;
            tableSetStat.TableBorderSetting = new TableBorderSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            tableSetStat.TableBorderInsideSetting = new TableBorderInsideSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            TableModel tableMdlStat = new TableModel(hTblStat, tableSetStat, new List<List<object>>());

            if (_exeption == null)
                return tableMdlStat;

            foreach (int r in _exeption)
            {
                double avgCor = CalcAveragePortfolioCorrelation( r, corelations );
                tableMdlStat.TableData.Add(new List<object>()
                {
                    r, Math.Round(avgCor, 2)
                });
            }
            return tableMdlStat;
        }
        /// <summary>
        /// Рассчет весовых коэф. и формирование таблицы 
        /// </summary>
        TableModel WeightTable(List<PortfelResultTest> results)
        {
            List<int> robotNum = new List<int>();
            List<double> volaty = new List<double>();
            List<double> weight = new List<double>();
            List<double> sample = new List<double>();

            foreach (PortfelResultTest res in results)
            {
                if (_exeption.Contains(res.NumberRobot) || _noDeals.Contains(res.NumberRobot))
                    continue;
                robotNum.Add(res.NumberRobot);
                Statistic stat = res.Statistic;
                double eqMax = stat.EquityPoint.Max();
                double eqMin = stat.EquityPoint.Min();
                volaty.Add((eqMax - eqMin) * 100 / _entrySize );
            }
            double volatyMax = volaty.Max();

            for ( int i=0; i < volaty.Count; i++)
            {
                weight.Add(1/(volaty[i] / volatyMax));
            }
            double normSum = weight.Sum();
            
            for (int i = 0; i < weight.Count; i++)
            {
                weight[i] /= normSum;
                sample.Add(weight[i] * 100000);
            }

            HeaderTable hTblStat = new HeaderTable();
            hTblStat.Headers = new List<string> { " № Робота", "волатильность (%)","Вес", "На 100 000" };
            TableSetting tableSetStat = new TableSetting();
            tableSetStat.BodySetting.SettingText.TextAligment = Export.Enums.Aligment.Center;
            tableSetStat.TableBorderSetting = new TableBorderSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            tableSetStat.TableBorderInsideSetting = new TableBorderInsideSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            TableModel tableMdlStat = new TableModel(hTblStat, tableSetStat, new List<List<object>>());

            for(int i=0; i<robotNum.Count; i++)
            {
                tableMdlStat.TableData.Add(new List<object>()
                {
                    robotNum[i], Math.Round( volaty[i], 2), Math.Round( weight[i], 2), Math.Round( sample[i], 2)
                });
            }
            return tableMdlStat;
        }
        /// <summary>
        ///  Добавоение статистики в отчет по опр. роботу
        /// </summary>
        void AddToReport(PortfelResultTest res, List<CorelationModel> corelations )
        {
            Statistic stat = res.Statistic;
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

            string nameStrategy = "Стратегия: " + res.NameStategy + "  № робота: " + res.NumberRobot;
            string testPeriod = "Период тестирования: " + stat.StartTime.ToString() + " - ";
            testPeriod += stat.EndTime.ToString();
            string capital = "Стартовый капитал: " + stat.InitialCapital.ToString();
            string posSize = "размер позиции: "; //+ res.PosSize.ToString();
            string averageCor = "Коэф. корреляцици в среднем по портфелю: ";
            double avgCor = CalcAveragePortfolioCorrelation(res.NumberRobot, corelations);

            if(_exeption != null && avgCor != double.NaN)
            {
                averageCor += Math.Round( avgCor, 2).ToString();
                
                if(Math.Round(avgCor, 2) > _coreLimit )
                    _exeption.Add(res.NumberRobot);
            }
            string symbTF = res.Symbol + " " + res.TimeFrameType + " " + res.TimeFramePeriod.ToString() + "\n";

            _totalActions.Add(() => _totalPDF.AddText(new Text(nameStrategy, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(testPeriod, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(capital, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(posSize, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(averageCor, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(symbTF, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text("Показатели статистики\n", setTxtBold)));
            _totalActions.Add(() => _totalPDF.AddTable(StatisticTable(res)));
            _totalActions.Add(() => _totalPDF.AddText(new Text("Доходность за период\n", setTxtBold)));
            EquityLineChart(res, _totalPDF, _totalActions);
            _totalActions.Add(() => _totalPDF.AddNewPage());
        }
        
        #endregion

        public override void GetAttributesPortfel()
        {
            DesParamStratetgy.Version = "1.0.0.0";
            DesParamStratetgy.DateRelease = new DateTime(2023, 3, 20);
            DesParamStratetgy.DateChange = new DateTime(2023, 6, 13);
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "PortfolioStat";
        }
    }
}
