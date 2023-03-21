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
        public ClientReport? _totalPDF = null;
        public List<Action>? _totalActions = null;
        List<int>? _exeption = null; // Номера роботов подлежащие удалению из портфеля

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

                if ( res.TimeFrameType == "День")
                    set.date.Add(dt.AddDays(i));
                else
                    set.date.Add(dt.AddMinutes(i * res.TimeFramePeriod ));
            }
            data.Add(equity);
            actions.Add(() => rep.AddChart(new Chart(new Line(new LineETS(set, data)))));
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

            if( avgCor != double.NaN)
            {
                averageCor += Math.Round( avgCor, 2).ToString();
                
                if(Math.Round(avgCor, 2) > 0.5 )
                    _exeption.Add(res.NumberRobot);
            }
            string symbTF = res.Symbol + " " + res.TimeFrameType + " " + res.TimeFramePeriod.ToString() + "\n";
   
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

            _totalActions.Add(() => _totalPDF.AddText(new Text(nameStrategy, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(testPeriod, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(capital, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(posSize, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(averageCor, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text(symbTF, setTxt)));
            _totalActions.Add(() => _totalPDF.AddText(new Text("Показатели статистики\n", setTxtBold)));
            _totalActions.Add(() => _totalPDF.AddTable(tableMdlStat));
            _totalActions.Add(() => _totalPDF.AddText(new Text("Доходность за период\n", setTxtBold)));
            EquityLineChart(res, _totalPDF, _totalActions);
            _totalActions.Add(() => _totalPDF.AddNewPage());
        }
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

            foreach (PortfelResultTest res in results )
            {
                AddToReport(res, corelations);    
            }
            SettingText setTxtBold = new SettingText();
            setTxtBold.FontSize = 12;
            setTxtBold.FontName = "Arial";
            setTxtBold.TextAligment = Export.Enums.Aligment.Left;
            setTxtBold.Bold = true;

            _totalActions.Add(() => _totalPDF.AddText(new Text("Рекомендуется удалить из портфеля элементы со следующими номерами:\n", setTxtBold)));
            _totalActions.Add(() => _totalPDF.AddTable(ExeptionTable(corelations)));
            _totalPDF.GenerateReport(_totalActions);
            _totalPDF.SaveDocument();
            return new List<List<int>>();
        }
        #endregion

        public override void GetAttributesPortfel()
        {
            DesParamStratetgy.Version = "1.0.0.0";
            DesParamStratetgy.DateRelease = "20.03.2023";
            DesParamStratetgy.DateChange = "20.03.2023";
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "PortfolioStat";
        }
    }
}
