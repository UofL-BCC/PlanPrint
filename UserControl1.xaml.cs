using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Threading;
using OxyPlot.Wpf;
using OxyPlot.Series;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Legends;
using PdfSharpCore.Pdf;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore;
using MigraDoc.DocumentObjectModel;
using Section = MigraDoc.DocumentObjectModel.Section;
using Colors = MigraDoc.DocumentObjectModel.Colors;
using MigraDoc.DocumentObjectModel.Tables;
using Paragraph = MigraDoc.DocumentObjectModel.Paragraph;
using Table = MigraDoc.DocumentObjectModel.Tables.Table;
using MigraDoc.Rendering;

namespace ConsoleApp1
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class UserControl1 : UserControl
    {

        string printerName = "";
        string patientName = "NA";
        string patientId = "NA";
        string courseId = "NA";
        string planId = "NA";
        string user = "NA";
        string printDate = "NA";
        Bitmap captImage;

        int count = 0;


        public UserControl1(ScriptContext context, Window window)
        {
            var plan = context.PlanSetup;
            var planSum = context.PlanSumsInScope.FirstOrDefault();
            var patient = context.Patient;
            var course = context.Course;

            //string printerName = "";
            //foreach (string s in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            //{
            //    if (s.Contains("eDoc"))
            //        printerName = s;
            //}

            //check pt not null
            if (patient == null)
            {
                MessageBox.Show("No patient loaded!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            //canonly have 1 plansum
            if (context.PlanSumsInScope.Count() > 1)
            {
                MessageBox.Show("Two or more PlanSums are loaded in Scope.\nPlease close the unused PlanSum.!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }


            string patientName = patient.LastName + " " + patient.FirstName;
            string patientId = patient.Id;
            string user = context.CurrentUser.Name;
            string printDate = DateTime.Now.ToString();


            var selectedPlanningItem = plan != null ? (PlanningItem)plan : (PlanningItem)planSum;

            if (course != null)
            {
                courseId = course.Id;
            }

            if (selectedPlanningItem != null)
            {
                planId = selectedPlanningItem.Id;
            }



            //now use Oxyplot to build a plot
            //window.Title = "DVH";
            //define the plotView
            PlotView plotView = new PlotView();
            //define the plotModel
            OxyPlot.PlotModel plotModel = new OxyPlot.PlotModel();
            plotModel.Axes.Add(new LinearAxis { Title = "Dose [cGy]", Position = AxisPosition.Bottom, MajorGridlineStyle= LineStyle.Solid });
            plotModel.Axes.Add(new LinearAxis { Title = "Ratio of Total Structure Volume [%]", Position = AxisPosition.Left, MinorGridlineStyle = LineStyle.Solid});
            FormatLegend(plotModel);
            //get the selected DVH structures
            var selectedStructures = plan.StructuresSelectedForDvh;


            //Plot the structures on the model
            PlotStructures(plan, selectedStructures, plotModel);

            //assign the model to the view
            plotView.Model = plotModel;

            //show the window, allow user to modify DVH structures they want to show?
            window.Content = plotView;



            //when the DVH window is closed, get a screenshot of the plan in view
            window.Closed += Window_Closed;

            //for now, close the script loading window by making a messagebox show, then run the screenshot thread
            MessageBox.Show("begin");


            //ScreenshotThreeView();
            //Thread.Sleep(5000);




            string assemblyLoc = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);




            var saveLocation = Path.Combine(assemblyLoc, "report.pdf");
            ExportToPDF(saveLocation, context);


        }

        public void ScreenshotThreeView()
        {

            //need to run this on a seperate thread so it can wait while the loading bar closes
            try
            {
                Thread trd = new Thread(new ThreadStart(this.ThreadTask));
                trd.IsBackground = true;
                trd.Start();
            }
            catch (Exception d)
            {

                MessageBox.Show(d.Message);
            }

        }

        public void ExportToPDF(string pdfFile, ScriptContext context)
        {
            var PatientId = context.Patient.Id;
            var PatientName = context.Patient.Name;
            var DOB = context.Patient.DateOfBirth.ToString();
            var PlanId = context.PlanSetup.Id;
            var CourseId = context.Course.Id;
            var Approval = context.PlanSetup.ApprovalStatus.ToString();
            var Modification = context.PlanSetup.HistoryUserName.ToString();
            var Date = context.PlanSetup.HistoryDateTime.ToString();
            var plan = context.PlanSetup;

            string assemblyLoc = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string threeViewImageLoc = Path.Combine(assemblyLoc, "output.png");




            Document migraDoc = new Document();
            Section section = migraDoc.AddSection();
            MigraDoc.DocumentObjectModel.Shapes.Image threeViewImage = section.AddImage(threeViewImageLoc);
            threeViewImage.Width = Unit.FromCentimeter(10);
            threeViewImage.Height = Unit.FromCentimeter(10);
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
            //Paragraph paragraph = section.AddParagraph();
            MigraDoc.DocumentObjectModel.Tables.Table table = new MigraDoc.DocumentObjectModel.Tables.Table();
            table.Borders.Width = 1;
            table.Borders.Color = Colors.White;
            table.AddColumn(Unit.FromCentimeter(6));
            table.AddColumn(Unit.FromCentimeter(6));
            Row row = table.AddRow();
            Cell cell = row.Cells[0];
            cell.AddParagraph("Patient ID:");
            cell = row.Cells[1];
            Paragraph paragraph = cell.AddParagraph();
            paragraph.AddFormattedText(PatientId, TextFormat.Bold);
            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Patient Name:");
            cell = row.Cells[1];
            paragraph = cell.AddParagraph();
            paragraph.AddFormattedText(PatientName, TextFormat.Bold);
            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Date of Birth:");
            cell = row.Cells[1];
            if (!string.IsNullOrEmpty(DOB))
            {
                cell.AddParagraph(DOB);
            }
            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Course ID:");
            cell = row.Cells[1];
            cell.AddParagraph(CourseId);
            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Plan ID:");
            cell = row.Cells[1];
            cell.AddParagraph(PlanId);
            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Approval Status:");
            cell = row.Cells[1];
            cell.AddParagraph(Approval);
            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Modification by:");
            cell = row.Cells[1];
            cell.AddParagraph(Modification);
            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("Modification Date/Time:");
            cell = row.Cells[1];
            cell.AddParagraph(Date);

            section.Add(table);

            Paragraph paragraph2 = section.AddParagraph("\n\n");
            paragraph2.AddFormattedText("Plan Summary", TextFormat.Bold);

            table = new Table();
            table.Borders.Width = 1;
            table.Borders.Color = Colors.Olive;
            for (int c = 0; c < 10; c++)
            {
                table.AddColumn(Unit.FromCentimeter(2.6));
            }
            row = table.AddRow();
            row.Shading.Color = Colors.PaleGoldenrod;
            cell = row.Cells[0];
            cell.AddParagraph("Course ID");
            cell = row.Cells[1];
            cell.AddParagraph("Plan ID");
            cell = row.Cells[2];
            cell.AddParagraph("Plan Type");
            cell = row.Cells[3];
            cell.AddParagraph("Target Volume ID");
            cell = row.Cells[4];
            cell.AddParagraph("Target a/b");
            cell = row.Cells[5];
            cell.AddParagraph("Presc. %");
            cell = row.Cells[6];
            cell.AddParagraph("Dose / Fraction");
            cell = row.Cells[7];
            cell.AddParagraph("Number of Fractions");
            cell = row.Cells[8];
            cell.AddParagraph("Total Dose Planned");
            cell = row.Cells[9];
            cell.AddParagraph("EQD2 Dose Planned");

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph(plan.Course.Id);
            cell = row.Cells[1];
            cell.AddParagraph(plan.Id);
            cell = row.Cells[2];
            cell.AddParagraph(plan.PlanType.ToString());
            cell = row.Cells[3];
            cell.AddParagraph(plan.TargetVolumeID);
            cell = row.Cells[4];
            //cell.AddParagraph(plan.TargetAB);
            cell = row.Cells[5];
            cell.AddParagraph(plan.TreatmentPercentage.ToString());
            cell = row.Cells[6];
            cell.AddParagraph(plan.DosePerFraction.ToString());
            cell = row.Cells[7];
            cell.AddParagraph(plan.NumberOfFractions.ToString());
            cell = row.Cells[8];
            cell.AddParagraph(plan.TotalDose.ToString());
            cell = row.Cells[9];


            //foreach (var plan in Table)
            //{
            //    row = table.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph(plan.CourseID);
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.PlanID);
            //    cell = row.Cells[2];
            //    cell.AddParagraph(plan.PlanType);
            //    cell = row.Cells[3];
            //    cell.AddParagraph(plan.TargetVolumeID);
            //    cell = row.Cells[4];
            //    cell.AddParagraph(plan.TargetAB);
            //    cell = row.Cells[5];
            //    cell.AddParagraph(plan.PrescPercentage);
            //    cell = row.Cells[6];
            //    cell.AddParagraph(plan.DosePerFraction);
            //    cell = row.Cells[7];
            //    cell.AddParagraph(plan.NumberOfFractions);
            //    cell = row.Cells[8];
            //    cell.AddParagraph(plan.TotalDosePlanned);
            //    cell = row.Cells[9];
            //    cell.AddParagraph(plan.EQD2DosePlanned);
            //}
            section.Add(table);

            Paragraph paragraph3 = section.AddParagraph("\n");
            //paragraph3.AddFormattedText("Total Planned EQD2 Dose: " + TotalEQD2Dose);

            Paragraph paragraph4 = section.AddParagraph("\n\n");
            paragraph4.AddFormattedText("Plan Details", TextFormat.Bold);

            Table table2 = new Table();
            table2.Borders.Width = 1;
            table2.Borders.Color = Colors.White;
            for (int c = 0; c < 4; c++)
            {
                table2.AddColumn(Unit.FromCentimeter(6));
            }
            //foreach (var plan in Details)
            //{
            //    row = table2.AddRow();
            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    Paragraph p = cell.AddParagraph();
            //    p.AddFormattedText(plan.PlanID, TextFormat.Bold);
            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph("Course ID:");
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.CourseID);

            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph("Plan Type:");
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.PlanType);
            //    cell = row.Cells[2];
            //    cell.AddParagraph("Technique:");
            //    cell = row.Cells[3];
            //    cell.AddParagraph(plan.Technique);

            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph("Dose / Fraction:");
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.DosePerFraction);
            //    cell = row.Cells[2];
            //    cell.AddParagraph("Prescribed Percentage:");
            //    cell = row.Cells[3];
            //    cell.AddParagraph(plan.PrescPercentage);

            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph("Number of Fractions:");
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.NumberOfFractions);

            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph("Target Volume ID:");
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.TargetVolumeID);
            //    cell = row.Cells[2];
            //    cell.AddParagraph("Target Volume a/b:");
            //    cell = row.Cells[3];
            //    cell.AddParagraph(plan.TargetAB);

            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph("Total Dose Planned:");
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.TotalDosePlanned);
            //    cell = row.Cells[2];
            //    cell.AddParagraph("EQD2 Dose Planned:");
            //    cell = row.Cells[3];
            //    cell.AddParagraph(plan.EQD2DosePlanned);

            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph("Approval Status:");
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.Status);

            //    row = table2.AddRow();
            //    cell = row.Cells[0];
            //    cell.AddParagraph("Modification by:");
            //    cell = row.Cells[1];
            //    cell.AddParagraph(plan.Modified);
            //    cell = row.Cells[2];
            //    cell.AddParagraph("Modification Date/Time:");
            //    cell = row.Cells[3];
            //    cell.AddParagraph(plan.Date);
            //}
            section.Add(table2);

            

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfSharp.Pdf.PdfFontEmbedding.None);
            pdfRenderer.Document = migraDoc;
            pdfRenderer.RenderDocument();
            pdfRenderer.PdfDocument.Save(pdfFile);
        }

        public void FormatLegend(PlotModel plotModel)
        {
           
           
            Legend legend = new Legend();
            legend.LegendOrientation = LegendOrientation.Vertical;
            legend.LegendPosition = LegendPosition.BottomLeft;
            legend.LegendPlacement = LegendPlacement.Outside;
            legend.LegendColumnSpacing = 5;
            legend.LegendSymbolLength = 25;
            legend.LegendFont = "Courier New";

            legend.LegendTitleFontWeight = OxyPlot.FontWeights.Bold;
            legend.LegendTitle = "DVH";
            legend.LegendTitle = legend.LegendTitle.PadRight(15);
            legend.LegendTitle += "Structure";
            legend.LegendTitle = legend.LegendTitle.PadRight(60);
            legend.LegendTitle += "Structure Status";
            legend.LegendTitle = legend.LegendTitle.PadRight(90);
            legend.LegendTitle += "Coverage[%/%]";
            legend.LegendTitle = legend.LegendTitle.PadRight(128);
            legend.LegendTitle += "Volume";
            legend.LegendTitle = legend.LegendTitle.PadRight(166);
            legend.LegendTitle += "Min Dose";
            int padCount = 166;
            padCount += 38;
            legend.LegendTitle = legend.LegendTitle.PadRight(padCount);
            legend.LegendTitle += "Max Dose";
            padCount += 38;
            legend.LegendTitle = legend.LegendTitle.PadRight(padCount);
            legend.LegendTitle += "Mean Dose";
            padCount += 38;
            //ignore these other stats for now
            //legend.LegendTitle = legend.LegendTitle.PadRight(padCount);
            //legend.LegendTitle += "Modal Dose";
            //padCount += 38;
            //legend.LegendTitle = legend.LegendTitle.PadRight(padCount);
            //legend.LegendTitle += "Median Dose";
            //padCount += 38;
            //legend.LegendTitle = legend.LegendTitle.PadRight(padCount);
            //legend.LegendTitle += "Std Dev";
            //padCount += 38;
            //legend.LegendTitle = legend.LegendTitle.PadRight(padCount);


            plotModel.Legends.Add(legend);



        }

        private void Window_Closed(object sender, EventArgs e)
        {
            try
            {
                Thread trd = new Thread(new ThreadStart(this.ThreadTask));
                trd.IsBackground = true;
                trd.Start();
            }
            catch (Exception d)
            {

                MessageBox.Show(d.Message);
            }
           
            
        }

        public void PlotStructures(PlanSetup plan, IEnumerable<Structure> structures, OxyPlot.PlotModel model)
        {

            foreach (var structure in structures)
            {
                //Make an invisible series for the legend header


                //get the DVH data

                DVHData dvh;
                if (plan.GetDVHCumulativeData(structure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.01) != null)
                {
                    dvh = plan.GetDVHCumulativeData(structure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.01);
                    
                    //make the table headers for the legend
                    string seriesTitle = StringBuilder(structure, dvh, plan);
                    count++;
                    //define a line series for the curve
                    var series = new LineSeries { Title = seriesTitle, Tag = structure.Id};

                    //define points from the dvh data
                    var points = new List<DataPoint>();
                    foreach (var dvhPoint in dvh.CurveData)
                    {
                        points.Add(new DataPoint(dvhPoint.DoseValue.Dose, dvhPoint.Volume));
                    }

                    //add the points to the series
                    series.Points.AddRange(points);

                    //add the series to the model
                    model.Series.Add(series);
                }

            }
        }

        public string StringBuilder(Structure structure, DVHData dvh, PlanSetup plan)
        {
            double stringCount = structure.Id.Length;

            string legendLabel = structure.Id;

            legendLabel = legendLabel.PadRight(30);


            if (structure.IsApproved == true)
            {
                legendLabel += "Approved";

            }
            else
            {
                legendLabel += "Unapproved";
            }

            legendLabel = legendLabel.PadRight(50);


            double doseCoverage = Math.Round(dvh.Coverage * 100);
            double samplingCoverage = Math.Round(dvh.SamplingCoverage*100);

            legendLabel += String.Format("{0}/{1}", doseCoverage, samplingCoverage);

            legendLabel = legendLabel.PadRight(70);


            legendLabel += String.Format("{0} cm\u00B3", Math.Round(dvh.Volume,1));

            legendLabel = legendLabel.PadRight(90);

            legendLabel += String.Format("{0}%", Math.Round((dvh.MinDose.Dose/plan.TotalDose.Dose) * 100,1));

            legendLabel = legendLabel.PadRight(110);

            legendLabel += String.Format("{0}%", Math.Round((dvh.MaxDose.Dose / plan.TotalDose.Dose) * 100, 1));

            legendLabel = legendLabel.PadRight(130);

            legendLabel += String.Format("{0}%", Math.Round((dvh.MeanDose.Dose / plan.TotalDose.Dose) * 100, 1));

            legendLabel = legendLabel.PadRight(150);


            //modal dose a little tricky
            //it is the dose value of which the largest volume receives (you can see this is the DVH differential graph)
            //but cant pull info straight form the graph
            //not just the most common dose value

            //find the median dose
            //ignore the last few stats prob dont need them
            //var doseList = dvh.CurveData.Select(c => Math.Round(c.DoseValue.Dose,1)).ToList();
            //var orderedList = doseList.OrderBy(c => c).ToList();
            //int middleIndex = orderedList.Count / 2;
            //double averagedValue;
            //if (orderedList.Count % 2 == 0)
            //{
            //    double middleValue1 = orderedList[middleIndex - 1];
            //    double middleValue2 = orderedList[middleIndex];
            //    averagedValue = (middleValue1 + middleValue2) / 2;
            //}
            //else
            //{
            //    averagedValue = orderedList[middleIndex];
            //}

            //var groupedList = doseList.GroupBy(c => c);
            //var maxFrequencyValue = groupedList.Max(g => g.Count());
            //var doseWithMaxFrequency = groupedList.First(c => c.Count() == maxFrequencyValue).Key;


            //legendLabel += String.Format("{0}%", Math.Round((averagedValue / plan.TotalDose.Dose) * 100, 1));



            return legendLabel;
        }

        public static IntPtr GetBackGroundWindow()
        {
            IntPtr foreground = GetForegroundWindow();
            IntPtr background = GetWindow(foreground, GW_HWNDPREV);
            return background;
        }

        private void ThreadTask()
        {
            Rectangle rectangle = new Rectangle();
            RECT rect;

            Thread.Sleep(3000);

            IntPtr activeWindow = GetForegroundWindow();
            //not working
            //IntPtr activeWindow = GetBackGroundWindow();


            GUITHREADINFO gUITHREADINFO = new GUITHREADINFO();
            //active = gUITHREADINFO.hwndactive;
            //active = gUITHREADINFO.hwndactive;
            //active = gUITHREADINFO.hwndFocus;
            //active = gUITHREADINFO.hwndCapture;
            //active = gUITHREADINFO.hwndMenuOwner;


            GetWindowRect(activeWindow, out rect);
            rectangle = new Rectangle(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);

            captImage = new Bitmap(rectangle.Width, rectangle.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);





            using (captImage)
            {

                // Get screen capture using CopyFromScreen method.
                using (Graphics g = Graphics.FromImage(captImage))
                {
                    g.CopyFromScreen(rectangle.X, rectangle.Y, 0, 0, rectangle.Size, CopyPixelOperation.SourceCopy);

                    float zoom = 1;
                    float padding = 50;

                    if (captImage.Width > g.VisibleClipBounds.Width)
                    {
                        zoom = g.VisibleClipBounds.Width /
                            captImage.Width;
                    }

                    if ((captImage.Height + padding) * zoom >
                            g.VisibleClipBounds.Height)
                    {
                        zoom = g.VisibleClipBounds.Height /
                            (captImage.Height + padding);
                    }

                    //already copying the image above,  dont need to draw it
                    //g.DrawImage(captImage, 0, padding,
                    //                                   captImage.Width * zoom,
                    //                                   captImage.Height * zoom);

                    System.Drawing.Font font = new System.Drawing.Font("Arial Unicode MS", 10.5f);
                    System.Drawing.Brush brush = new SolidBrush(System.Drawing.Color.Black);

                    g.DrawString(string.Format("Patient ID: {0}, Patient Name: {1}", patientId, patientName), font, brush, new PointF(10, 10));
                    g.DrawString(string.Format("Course ID: {0}, Plan ID: {1},", courseId, planId), font, brush, new PointF(10, 30));
                    string outText = string.Format("(ESAPI ScreenCapture) Printed on {0} by {1}", printDate, user);
                    SizeF stringSize = g.MeasureString(outText, font, 1000);
                    g.DrawString(outText, font, brush, new PointF(g.VisibleClipBounds.Width - stringSize.Width - 10, 10));


                }
                string assemblyLoc = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                //MessageBox.Show(assemblyLoc);
                captImage.Save(Path.Combine(assemblyLoc, "output.png"));






            }
        }

        /// <summary>
        /// Define RECT class
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        /// <summary>
        /// Gets the size of the bounding rectangle of the specified window.
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="rect"></param>
        /// <returns></returns>
        [DllImport("User32.Dll")]
        static extern int GetWindowRect(IntPtr hWnd, out RECT rect);

        /// <summary>
        /// Take a screen capture and print.
        /// </summary>
        /// <returns></returns>
        [DllImport("user32.dll")]
        extern static IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
        private const uint GW_HWNDPREV = 3;


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

  


        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public uint flags;
            public IntPtr hwndactive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;

        }


        //private void ThreadTask()
        //{

        //    Rectangle rectangle = new Rectangle();

        //    // Wait 500 miliseconds for script progress bar to disappear
        //    Thread.Sleep(500);

        //    RECT rect;
        //    IntPtr active = GetForegroundWindow();
        //    GetWindowRect(active, out rect);
        //    rectangle = new Rectangle(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);

        //    captImage = new Bitmap(rectangle.Width, rectangle.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

        //    // Get screen capture using CopyFromScreen method.
        //    using (Graphics g = Graphics.FromImage(captImage))
        //    {
        //        g.CopyFromScreen(rectangle.X, rectangle.Y, 0, 0, rectangle.Size, CopyPixelOperation.SourceCopy);
        //    }

        //    PrintDocument pd = new PrintDocument();

        //    pd.PrintPage += new PrintPageEventHandler(pd_PrintPage);
        //    pd.DefaultPageSettings.Landscape = true;
        //    if (printerName != "")
        //        pd.PrinterSettings.PrinterName = printerName;

        //    for (var index = 0; index < pd.PrinterSettings.PaperSizes.Count; index++)
        //    {
        //        if (pd.PrinterSettings.PaperSizes[index].PaperName.Contains("A3") == true)
        //        {
        //            pd.DefaultPageSettings.PaperSize = pd.PrinterSettings.PaperSizes[index];
        //            break;
        //        }
        //    }

        //    pd.PrinterSettings.PrintToFile = true;

        //    pd.Print();

        //    MessageBox.Show("Done.");
        //}

        //private void pd_PrintPage(object sender, PrintPageEventArgs e)
        //{
        //    float zoom = 1;
        //    float padding = 50;

        //    if (captImage.Width > e.Graphics.VisibleClipBounds.Width)
        //    {
        //        zoom = e.Graphics.VisibleClipBounds.Width /
        //            captImage.Width;
        //    }

        //    if ((captImage.Height + padding) * zoom >
        //            e.Graphics.VisibleClipBounds.Height)
        //    {
        //        zoom = e.Graphics.VisibleClipBounds.Height /
        //            (captImage.Height + padding);
        //    }

        //    e.Graphics.DrawImage(captImage, 0, padding,
        //                                       captImage.Width * zoom,
        //                                       captImage.Height * zoom);

        //    Font font = new Font("Arial Unicode MS", 10.5f);
        //    System.Drawing.Brush brush = new SolidBrush(System.Drawing.Color.Black);

        //    e.Graphics.DrawString(string.Format("Patient ID: {0}, Patient Name: {1}", patientId, patientName), font, brush, new PointF(10, 10));
        //    e.Graphics.DrawString(string.Format("Course ID: {0}, Plan ID: {1},", courseId, planId), font, brush, new PointF(10, 30));
        //    string outText = string.Format("(ESAPI ScreenCapture) Printed on {0} by {1}", printDate, user);
        //    SizeF stringSize = e.Graphics.MeasureString(outText, font, 1000);
        //    e.Graphics.DrawString(outText, font, brush, new PointF(e.Graphics.VisibleClipBounds.Width - stringSize.Width - 10, 10));
        //}




    }
    //not currently using
    //public class ScreenCapture
    //{
    //    [DllImport("user32.dll")]
    //    private static extern IntPtr GetForegroundWindow();

    //    [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
    //    public static extern IntPtr GetDesktopWindow();

    //    [StructLayout(LayoutKind.Sequential)]
    //    private struct Rect
    //    {
    //        public int Left;
    //        public int Top;
    //        public int Right;
    //        public int Bottom;
    //    }

    //    [DllImport("user32.dll")]
    //    private static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);

    //    public static /*System.Windows.Controls.Image*/Bitmap CaptureDesktop()
    //    {
    //        var window = CaptureWindow(GetDesktopWindow());
    //        return window;
    //    }

    //    public Bitmap CaptureActiveWindow()
    //    {
    //        return CaptureWindow(GetForegroundWindow());
    //    }

    //    public static Bitmap CaptureWindow(IntPtr handle)
    //    {
    //        var rect = new Rect();
    //        GetWindowRect(handle, ref rect);
    //        var bounds = new System.Drawing.Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
    //        var result = new Bitmap(bounds.Width, bounds.Height);

    //        using (var graphics = Graphics.FromImage(result))
    //        {
    //            graphics.CopyFromScreen(new System.Drawing.Point(bounds.Left, bounds.Top), System.Drawing.Point.Empty, bounds.Size);
    //        }

    //        return result;
    //    }
    //}

   
    






}
