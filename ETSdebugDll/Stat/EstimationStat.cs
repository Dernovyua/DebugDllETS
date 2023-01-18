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

namespace ETSdebugDll
{
    public class UserParam
    {
        string  name = "";
        double  value = 0;
        int     digits = 6;
    }
    public class EstimationStat
    {
        public double profit = 0; // Профит
        public double drawDown = 0; // Максимальная просадка
        public double recoveryFactor = 0; // Факттор восстановления
        public double averageDeal = 0; // Средняя прибыль на сделку
        public double dealCount = 0; // Количество сделок 
        List<UserParam> userParam;
    };

    public class StatisticDebugScript : StatisticScript
    {
        int _passCount = 0;
        List<EstimationStat> _estStat;

        public override void SetUserStatisticParamOnEndTest(Statistic stat)
        {
            if (_passCount == 0)
                _estStat = new List<EstimationStat>();
            
            if( stat.FactorRecovery >= 2 )
            {
                EstimationStat est = new EstimationStat();
                est.profit = stat.NetProfitLoss;
                est.averageDeal = stat.AvaregeProfit;
                est.recoveryFactor = stat.FactorRecovery;
                est.dealCount = stat.TradeDealsList.Count;
                est.drawDown = stat.MaxDrownDown;
                _estStat.Add(est);
            }
        }

        void SaveResultPDF( string fNameFull )
        {
            iTextSharp.text.Document doc = new iTextSharp.text.Document();
            PdfWriter.GetInstance(doc, new FileStream("pdfTables.pdf", FileMode.Create));
            doc.Open();
            //BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            //iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);
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

            doc.Close();
 
        }
    void SaveResultCSV(string fNameFull, string delimiter )
        {

        }

        public override void GetAttributes()
        {
            DesParamStratetgy.Version = "2";
            DesParamStratetgy.DateRelease = "16.01.2023";
            DesParamStratetgy.DateChange = "17.01.2023";
            DesParamStratetgy.Description = "вариант 1, смотрим";
            DesParamStratetgy.Change = "";
            DesParamStratetgy.NameStrategy = "EstimationStat";
        }
    }
}
