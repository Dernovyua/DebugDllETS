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

namespace ETSdebugDll.PortfolioBuilder
{
    public class PortfolioBuilder : PortfelScript
    {
        public ClientReport? _totalPDF = null;
        public List<Action>? _totalActions = null;
        double _entrySize = 10000;
        int _pornfolioNumRobot = 5;
        double _coreLimit = 0.5;

        List<List<PortfelResultTest>> _portfolioSet = new List<List<PortfelResultTest>>();
        List<PortfelResultTest> _sourceCopy = new List<PortfelResultTest>();

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
            _sourceCopy = new List<PortfelResultTest>();
            _portfolioSet = new List<List<PortfelResultTest>>();

            foreach ( PortfelResultTest res in results)
                _sourceCopy.Add(res);
            
            while( _sourceCopy.Count > _pornfolioNumRobot )
                BuildPortfolio(corelations);
            
            string path = PathSaveResult;
            string name = "\\PortfolioBuilder.pdf";
            _totalPDF = new ClientReport();
            _totalPDF.SetExport(new Pdf(path, name, false));
            _totalActions = new List<Action>();


            _totalPDF.GenerateReport(_totalActions);
            _totalPDF.SaveDocument();
            return new List<List<int>>();
        }
        /// <summary>
        /// Проверяем условия присутствия элемента в указанном портфеле
        /// </summary>
        bool CheckCorrelation( int robotNum, List<PortfelResultTest> portfolio, List<CorelationModel> corelations)
        {
            CorelationModel? corMdl = null;

            // Находим моель корреляции для указанного робота
            foreach(CorelationModel c in corelations)
            {
                if( c.NumberRobot == robotNum )
                {
                    corMdl = c;
                    break;
                }
            }

            if(corMdl == null)
                return false;
            
            double portfolioCore = 0;
            int findeCount = 0;

            // Смотрим как указанный робот коррелирует с существ. порфелем
            foreach( PortfelResultTest p in portfolio )
            {
                foreach( CorValueModel cv in corMdl.CorValues )
                {
                    if( p.NumberRobot == cv.NumberRobot )
                    {
                        findeCount++;
                        portfolioCore += cv.Corelation;
                    }
                }
            }

            if (findeCount == 0 || findeCount != portfolio.Count)
                return false;

            portfolioCore /= (double)findeCount;

            if (portfolioCore <= _coreLimit)
                return true;

            return false;
        }
        /// <summary>
        /// удаляем элемент из списка роботов
        /// </summary>
        void DeleteFromSource( int robotNum )
        {
            for (int i = 0; i < _sourceCopy.Count; i++)
            {
                if (robotNum == _sourceCopy[i].NumberRobot)
                { 
                    _sourceCopy.RemoveAt(i);
                    return;
                }
            }
        }
        /// <summary>
        /// Формируем портфель
        /// </summary>
        void BuildPortfolio(List<CorelationModel> corelations)
        {
            List<PortfelResultTest> portfolio = new List<PortfelResultTest>();

            foreach (PortfelResultTest res in _sourceCopy)
            {
                if (portfolio.Count == 0 || CheckCorrelation(res.NumberRobot, portfolio, corelations))
                {
                    portfolio.Add(res);
                }
                if( portfolio.Count >= _pornfolioNumRobot)
                    break;
            }

            if (portfolio.Count > 0)
            {
                foreach (PortfelResultTest res in portfolio)
                {
                    DeleteFromSource(res.NumberRobot);
                }
                _portfolioSet.Add(portfolio);
            }
        }


        TableModel WeightTable(List<PortfelResultTest> results)
        {
            List<int> robotNum = new List<int>();
            List<double> volaty = new List<double>();
            List<double> weight = new List<double>();
            List<double> sample = new List<double>();

            foreach (PortfelResultTest res in results)
            {
                robotNum.Add(res.NumberRobot);
                Statistic stat = res.Statistic;
                double eqMax = stat.EquityPoint.Max();
                double eqMin = stat.EquityPoint.Min();
                volaty.Add((eqMax - eqMin) * 100 / _entrySize);
            }
            double volatyMax = volaty.Max();

            for (int i = 0; i < volaty.Count; i++)
            {
                weight.Add(1 / (volaty[i] / volatyMax));
            }
            double normSum = weight.Sum();

            for (int i = 0; i < weight.Count; i++)
            {
                weight[i] /= normSum;
                sample.Add(weight[i] * 100000);
            }

            HeaderTable hTblStat = new HeaderTable();
            hTblStat.Headers = new List<string> { " № Робота", "волатильность (%)", "Вес", "На 100 000" };
            TableSetting tableSetStat = new TableSetting();
            tableSetStat.BodySetting.SettingText.TextAligment = Export.Enums.Aligment.Center;
            tableSetStat.TableBorderSetting = new TableBorderSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            tableSetStat.TableBorderInsideSetting = new TableBorderInsideSetting() { BorderLineStyle = SettingBorderLineStyle.None };
            TableModel tableMdlStat = new TableModel(hTblStat, tableSetStat, new List<List<object>>());

            for (int i = 0; i < robotNum.Count; i++)
            {
                tableMdlStat.TableData.Add(new List<object>()
                {
                    robotNum[i], Math.Round( volaty[i], 2), Math.Round( weight[i], 2), Math.Round( sample[i], 2)
                });
            }
            return tableMdlStat;
        }

        #endregion

        public override void GetAttributesPortfel()
        {
            DesParamStratetgy.Version = "1.0.0.1";
            DesParamStratetgy.DateRelease = "23.03.2023";
            DesParamStratetgy.DateChange = "23.03.2023";
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "PortfolioBuilder";
        }
    }
}

