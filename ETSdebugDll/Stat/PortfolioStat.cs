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
using System.Runtime.CompilerServices;
using System.Windows.Media.Animation;
using ScriptSolution.Model.Portfels.PortfelTest;

namespace PortfolioStat
{
    internal class PortfolioStat : PortfelScript
    {     

        public override List<List<int>> ExecuteEndOptimization( List< PortfelResultTest> prt, List<CorelationModel> cmdl )
        {
            List<List<int>> res = new List<List<int>>();
            return res;
        }
    }
}
