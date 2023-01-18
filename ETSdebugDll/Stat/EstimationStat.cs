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

namespace EstimationStat
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
        public List<UserParam> userParam;// Параметры оптимизации
    };

    public class StatisticDebugScript : StatisticScript
    {
        int _passCount = 0;
        List<EstimationStat> _estStat;
        string[] _statNames = { "Profit", "DD", "Recovery", "Avg. Deal", "Deal Count" };

        public override void SetUserStatisticParamOnEndTest(Statistic stat)
        {
            if (_passCount == 0)
                _estStat = new List<EstimationStat>();

            var oParams = OptimParams;
            var uStat = UserStatistics;

            if( stat.FactorRecovery >= uStat[0].Value )
            {
                EstimationStat est = new EstimationStat();
                est.userParam = new List<UserParam>();

                est.profit = stat.NetProfitLoss;
                est.averageDeal = stat.AvaregeProfit;
                est.recoveryFactor = stat.FactorRecovery;
                est.dealCount = stat.TradeDealsList.Count;
                est.drawDown = stat.MaxDrownDown;
                
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

            if (_passCount == 10)
            {
                SaveResultCSV("C:\\DOCS\\out.csv", ";");
                //SaveResultPDF("C:\\DOCS\\out.pdf");
            }
        }

        #region Сохранение результатов в PDF 
        
        void SaveResultPDF( string fNameFull )
        {
            var document = new iTextSharp.text.Document();
            using (var writer = PdfWriter.GetInstance(document, new FileStream( fNameFull, FileMode.Create)))
            {
                document.Open();
                PdfPTable table = new PdfPTable(5);
                // Добавляем заголовки 
                PdfPCell cell = new PdfPCell(new Phrase(new Phrase("Profit")));
                cell.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
                table.AddCell(cell);

                cell = new PdfPCell(new Phrase(new Phrase("Max DD")));
                cell.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
                table.AddCell(cell);

                cell = new PdfPCell(new Phrase(new Phrase("Recovery")));
                cell.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
                table.AddCell(cell);

                cell = new PdfPCell(new Phrase(new Phrase("Avg. Deal")));
                cell.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
                table.AddCell(cell);

                cell = new PdfPCell(new Phrase(new Phrase("Deal Count")));
                cell.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
                table.AddCell(cell);

                for (int i = 0; i < _estStat.Count; i++)
                {
                    table.AddCell(new Phrase(_estStat[i].profit.ToString()));
                    table.AddCell(new Phrase(_estStat[i].drawDown.ToString()));
                    table.AddCell(new Phrase(_estStat[i].recoveryFactor.ToString()));
                    table.AddCell(new Phrase(_estStat[i].averageDeal.ToString()));
                    table.AddCell(new Phrase(_estStat[i].dealCount.ToString()));
                }

                document.Add(table);

                document.Close();
                writer.Close();
            }
        }

        #endregion

        #region Сохранение результатов в CSV

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
                        str += _estStat[i].userParam[j].value.ToString();

                        if ( i < _estStat[i].userParam.Count - 1)
                            str += delimiter;
                    }
                    sw.WriteLine(str);
                }
                sw.Close();
            }
        }

        #endregion
        public override List<UserParamModel> CreateUserStatisticParam()
        {
            List<UserParamModel> list = new List<UserParamModel>();
            list.Add(new UserParamModel { Name = "Recovery >= ", IsOptimization = false });
            return list;
        }

        public override void GetAttributes()
        {
            DesParamStratetgy.Version = "1";
            DesParamStratetgy.DateRelease = "16.01.2023";
            DesParamStratetgy.DateChange = "17.01.2023";
            DesParamStratetgy.Description = "";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "EstimationStat";
        }
    }
}
