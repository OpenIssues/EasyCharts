using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2021.DocumentTasks;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;

namespace EasyCharts
{
    internal class Program
    {
        [MTAThread]
        static void Main(string[] args)
        {
            OpenFile1(@"漏斗图.pptx");

            //Console.WriteLine(Thread.CurrentThread.ManagedThreadId);

            //var thread = new Thread(() =>
            //{
            //    var application = new Application();
            //    application.Run(new MainWindow());
            //});

            //System.Threading.Tasks.Task.Run(() =>
            //{
            //    var application = new Application();
            //    application.Run(new MainWindow());
            //}).Wait();

            //var dateTime = new DateTime(1899, 12, 31);
            //var newDateTime = dateTime.AddDays(37261);

            //Console.WriteLine(newDateTime.ToString("yyyy/M/d"));
            //Console.WriteLine(newDateTime);

            //double a = 0.00000001;
            //Console.WriteLine(100 / a);

            //var type = typeof(int);
            //var type1 = typeof(int);
            //var type2 = typeof(bool);

            //Console.WriteLine(type.Equals(type1));
            //Console.WriteLine(type.Equals(type2));

            //Console.WriteLine(string.Format("hello{1} {0}", 100.0, 1000));

            //OpenFile(@"I:\Source\Project\pptx2enbx\PptxToEnbx.Tests\TestFiles\图表\图表 默认旭日图.pptx");
        }

        private static void OpenFile(string file)
        {
            var document = PresentationDocument.Open(file, false);

            var slideParts = document.PresentationPart?.SlideParts;

            var currentSlidePart = slideParts?.ToArray()[0];

            var slideCommonSlideData = currentSlidePart?.Slide.CommonSlideData;

            var shapeTree = slideCommonSlideData?.ShapeTree;

            var graphicFrame = shapeTree?.GetFirstChild<AlternateContent>();

            var graphicData = graphicFrame?.Descendants<GraphicData>()?.FirstOrDefault();

            var chartReference = graphicData?.GetFirstChild<ChartReference>();

            if (chartReference?.Id?.Value is null)
            {
                Debugger.Break();
                return;
            }
            var chartPart = (ChartPart)currentSlidePart.GetPartById(chartReference.Id.Value);
            var chartSpace = chartPart.ChartSpace;

            var chart = chartSpace.GetFirstChild<Chart>();

            var chartTitle = chart?.Title;
        }

        private static void OpenFile1(string file)
        {
            var document = PresentationDocument.Open(file, false);

            var slideParts = document.PresentationPart?.SlideParts;

            var currentSlidePart = slideParts?.ToArray()[0];

            var slideCommonSlideData = currentSlidePart?.Slide.CommonSlideData;

            var shapeTree = slideCommonSlideData?.ShapeTree;

            var graphicFrame = shapeTree?.GetFirstChild<AlternateContent>();

            var graphicData = graphicFrame?.Descendants<GraphicData>()?.FirstOrDefault();

            //var chartReference = graphicData?.GetFirstChild<ChartReference>();

            //if (chartReference?.Id?.Value is null)
            //{
            //    Debugger.Break();
            //    return;
            //}
            //var chartPart = (ChartPart)currentSlidePart.GetPartById(chartReference.Id.Value);
            //var chartSpace = chartPart.ChartSpace;

            //var chart = chartSpace.GetFirstChild<Chart>();

            //var chartTitle = chart?.Title;

            var chartRef = graphicData?.GetFirstChild<OpenXmlUnknownElement>();
            var id = chartRef.ExtendedAttributes.FirstOrDefault().Value;
            var part = (ExtendedChartPart)currentSlidePart.GetPartById(id);
            var series = part.ChartSpace.Chart.PlotArea.PlotAreaRegion.GetFirstChild<Series>();
            var layoutIdValue = series.LayoutId.Value;
        }
    }
}
