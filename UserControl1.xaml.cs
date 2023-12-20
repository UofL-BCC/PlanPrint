//using iText.IO.Image;
//using iTextSharp.text;
//using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
//using System.Windows.Shapes;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing.Printing;
using System.Diagnostics;

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




        public UserControl1(ScriptContext context /*Window window*/)
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


            



            Thread trd = new Thread(new ThreadStart(this.ThreadTask));
            trd.IsBackground = true;
            trd.Start();




        }

        private void ThreadTask()
        {
            Rectangle rectangle = new Rectangle();
            RECT rect;
            Thread.Sleep(1000);

            IntPtr activeWindow = GetForegroundWindow();

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

                    Font font = new Font("Arial Unicode MS", 10.5f);
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
