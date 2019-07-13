using RPAValidator.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using AForge.Imaging.Filters;
using System.Windows.Interop;

namespace RPAValidator.Views
{
    /// <summary>
    /// Lógica de interacción para FakeUI.xaml
    /// </summary>
    public partial class FakeUI : Window
    {
        private MainWindow _LauncherForm;

        private String _FolderPicFullPath = @"D:\Personal\TFG\RPAValidator\RPAValidator\Resources\img\";
        private String _FolderPicPath = @"\Resources\img\";
        private System.Drawing.Image _Pic;
        internal System.Drawing.Image Pic { get => _Pic; set => _Pic = value; }

        private List<Event> _ExpectedEvents;
        private List<Event> _SucessfulEvents;
        Event _CurrentExpectedEvent;
        Event _FailureEvent;
        private int _CurrentEventIndex;
        private float _ClickOffset = 30f;
        private float diffThreshold = 0.04f;
        private System.Windows.Point _PreviousClickPoint;

        String keystroke = "";

        bool borderFound;
        bool realBorderFound;

        BitmapImage bmImage;
        Bitmap _Image;
        
        public FakeUI(MainWindow iParent)
        {
            InitializeComponent();
            _LauncherForm = iParent;

            InitializeComponent();

            _CurrentEventIndex = 0;

            _ExpectedEvents = new List<Event>();
            _SucessfulEvents = new List<Event>();

            //fakeTextBox.Visible = false;
            //fakeTextBox.Enabled = false;
            //fakeTextBox.Height = 25;
            //fakeTextBox.Width = 250;
        }
        public void AddExpectedEvent(List<String> values)
        {
            _ExpectedEvents.Add(new Event(values));
        }
        public void Start()
        {
            Show();
            NextImage();
        }
        public void Restart()
        {
            _CurrentExpectedEvent = null;
            _CurrentEventIndex = 0;
            _SucessfulEvents = new List<Event>();
        }
        private void NextImage()
        {
            if (_CurrentEventIndex < _ExpectedEvents.Count)
            {
                _CurrentExpectedEvent = _ExpectedEvents[_CurrentEventIndex];

                String picPath = _FolderPicFullPath;
                bool isKeystroke = _CurrentExpectedEvent.IsKeystroke();

                // Si es una introducción de texto, debemos seguir imprimiendo la anterior captura de pantalla, pues la que pertenece al evento del keystroke ya tiene el texto introducido
                // En caso de que sea un click (caso contrario), si imprimiremos la captura de pantalla correspondiente.
                if (isKeystroke)
                    picPath += _SucessfulEvents.Last<Event>().PicPath;
                else
                    picPath += _CurrentExpectedEvent.PicPath;

                if (File.Exists(picPath))
                {
                    if(_Image != null)
                        _Image.Dispose();

                    bmImage = new BitmapImage();
                    bmImage.BeginInit();
                    bmImage.UriSource = new Uri(picPath, UriKind.Absolute);
                    bmImage.EndInit();
                    PicBox.Source = bmImage;

                    using (MemoryStream outStream = new MemoryStream())
                    {
                        BitmapEncoder enc = new BmpBitmapEncoder();
                        enc.Frames.Add(BitmapFrame.Create(bmImage));
                        enc.Save(outStream);
                        _Image = new System.Drawing.Bitmap(outStream);
                    }

                    if (isKeystroke)
                    {
                        FakeTextBox.IsEnabled = true;
                        keystroke = "";
                                                        
                        TextBoxTransformation(bmImage);

                        FakeTextBox.Focus();
                    }
                }
            }
            else
            {
                TestFinished(false);
                // Generar successful log
            }
        }
        private void PicBox_MouseDown(object sender, MouseButtonEventArgs e)
        {
            System.Windows.Point clickCoord = e.GetPosition(this);
            if (_CurrentExpectedEvent.IsCursor())
            {
                //TODO: Obtener las coordenadas de click sobre la imagen y calcular la distancia
                double distance = 0;

                distance = GetDistance(_CurrentExpectedEvent.Click_coord, clickCoord);

                //TODO: Comprobar (en tiempo de ejecución) que botón del ratón ha sido pulsado mirando en sus propiedades
                if (distance < _ClickOffset && _CurrentExpectedEvent.Event_info.ToUpper().Contains(e.ChangedButton.ToString().ToUpper()))
                {
                    //_PreviousClickPoint = _CurrentExpectedEvent.Click_coord;
                    _PreviousClickPoint = new System.Windows.Point(e.GetPosition(this).X, e.GetPosition(this).Y);
                    _SucessfulEvents.Add(_CurrentExpectedEvent);
                    _CurrentEventIndex++;

                    NextImage();
                }
                else
                {
                    List<String> values = new List<String>();
                    values.Add(""); // [0] Id
                                    //values.Add(readValues[3]); // [3] X coord
                                    //values.Add(readValues[4]); // [4] Y coord
                    values.Add(clickCoord.X.ToString());
                    values.Add(clickCoord.Y.ToString());
                    values.Add(Event.EventType.Cursor.ToString("g")); // [6] Event type
                    values.Add("{" + e.ChangedButton.ToString().ToUpper() + " MOUSE}" + clickCoord.X.ToString() + " - " + clickCoord.Y.ToString()); // [8] Event type information
                    values.Add(""); // [10] Image name
                    _FailureEvent = new Event(values);
                    TestFinished(true);
                }
            }
            else
            {
                // Si se espera un keystroke, pero se realiza un click en la imagen, entrará por aquí
                List<String> values = new List<String>();
                values.Add(""); // [0] Id
                                //values.Add(readValues[3]); // [3] X coord
                                //values.Add(readValues[4]); // [4] Y coord
                values.Add(clickCoord.X.ToString());
                values.Add(clickCoord.Y.ToString());
                values.Add(Event.EventType.Cursor.ToString("g")); // [6] Event type
                values.Add("{" + e.ChangedButton.ToString().ToUpper() + " MOUSE}" + clickCoord.X.ToString() + " - " + clickCoord.Y.ToString()); // [8] Event type information
                values.Add(""); // [10] Image name
                _FailureEvent = new Event(values);

                TestFinished(true);
            }
        }
        private static double GetDistance(System.Windows.Point point1, System.Windows.Point point2)
        {
            //pythagorean theorem c^2 = a^2 + b^2
            //thus c = square root(a^2 + b^2)
            double a = (double)(point2.X - point1.X);
            double b = (double)(point2.Y - point1.Y);

            return Math.Sqrt(a * a + b * b);
        }

        private bool validKey(Key k)
        {
            bool valid = ((k != Key.LeftAlt)
                            && (k != Key.RightAlt)
                            && (k != Key.LeftCtrl)
                            && (k != Key.RightCtrl)
                            && (k != Key.LeftShift)
                            && (k != Key.RightShift)
                            && (k != Key.Escape)
                            && (k != Key.System)
                            /*&& (k != Key.Back)*/);

            return valid;
        }

        private void FakeUIWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (_CurrentExpectedEvent.IsKeystroke())
            {
                if (validKey(e.Key))
                {
                    if (e.Key.Equals(Key.Back))
                    {
                        keystroke = keystroke.Remove(keystroke.Length - 1);
                    }
                    else
                    {
                        String key = e.Key.ToString();

                        if (e.Key == Key.Space)
                            key = " ";

                        keystroke += key;

                        if (keystroke.ToUpper().Equals(_CurrentExpectedEvent.Event_info.ToUpper()))
                        {
                            _SucessfulEvents.Add(_CurrentExpectedEvent);
                            _CurrentEventIndex++;
                            _CurrentExpectedEvent = _ExpectedEvents[_CurrentEventIndex];

                            String picPath = _FolderPicFullPath;
                            picPath += _CurrentExpectedEvent.PicPath;
                            BitmapImage bmImage = new BitmapImage();
                            bmImage.BeginInit();
                            bmImage.UriSource = new Uri(picPath, UriKind.Absolute);
                            bmImage.EndInit();
                            PicBox.Source = bmImage;

                            FakeTextBox.Text = "";
                            FakeTextBox.IsEnabled = false;
                            FakeTextBox.Visibility = Visibility.Hidden;
                        }
                    }
                }
            }
            else
            {
                // Si se espera un click, pero se realiza un tecleado, entrará por aquí
                List<String> values = new List<String>();
                values.Add(""); // [0] Id
                                //values.Add(readValues[3]); // [3] X coord
                                //values.Add(readValues[4]); // [4] Y coord
                values.Add("0");
                values.Add("0");
                values.Add(Event.EventType.Keystrokes.ToString("g")); // [6] Event type
                values.Add(e.Key.ToString()); // [8] Event type information
                values.Add(""); // [10] Image name

                _FailureEvent = new Event(values);

                TestFinished(true);
            }
        }
        private void TestFinished(bool interrupted = false)
        {
            GenerateReport(interrupted);
            _LauncherForm.TestFinished(interrupted);
        }
        private void GenerateReport(bool interrupted)
        {
            Excel.Application ApExcel = new Excel.Application();
            object opc = Type.Missing;

            Excel.Workbook libro;
            libro = ApExcel.Workbooks.Add(opc);

            ApExcel.Visible = true;

            libro = ApExcel.Workbooks.Add(opc);
            Excel.Worksheet hoja = new Excel.Worksheet();

            hoja = (Excel.Worksheet)libro.Sheets.Add(opc, opc, opc, opc);

            hoja.Activate();

            hoja.Cells[1, 1] = "Trace";

            int row = 1;
            foreach (Event e in _SucessfulEvents)
            {
                hoja.Cells[row, 2] = e.Event_type.ToString("g");
                hoja.Cells[row, 3] = e.Event_info;

                row++;
            }

            if (interrupted)
            {

                hoja.Cells[row, 3] = "Failure";

                row++;
                row++;
                hoja.Cells[row, 1] = "Cause";
                hoja.Cells[row, 2] = "Expected";
                hoja.Cells[row, 3] = _CurrentExpectedEvent.Event_type.ToString("g");
                hoja.Cells[row, 4] = _CurrentExpectedEvent.Event_info;

                row++;
                hoja.Cells[row, 2] = "Received";
                hoja.Cells[row, 3] = _FailureEvent.Event_type.ToString("g");
                hoja.Cells[row, 4] = _FailureEvent.Event_info;
            }
        }
        public float GetBrightness(System.Windows.Media.Color color)
        {
            float num = ((float)color.R) / 255f;
            float num2 = ((float)color.G) / 255f;
            float num3 = ((float)color.B) / 255f;
            float num4 = num;
            float num5 = num;
            if (num2 > num4)
                num4 = num2;
            if (num3 > num4)
                num4 = num3;
            if (num2 < num5)
                num5 = num2;
            if (num3 < num5)
                num5 = num3;
            return ((num4 + num5) / 2f);
        }
        private void TextBoxTransformation(BitmapImage bmImage)
        {
            System.Windows.Point initCoord = new System.Windows.Point(_PreviousClickPoint.X, _PreviousClickPoint.Y), startPoint;
            System.Windows.Media.Color c = GetCoordinateColor(/*bmImage, */_PreviousClickPoint.X, _PreviousClickPoint.Y), startPointColor;
            System.Windows.Media.Color foreColor;

            // Primer desplazamiento. Comienzo: @initCoord.
            double yPos = searchTopBorder(bmImage, initCoord/*, c*/);
            // Segundo desplazamiento. Comienzo: @startPoint (coordenada del píxel del límite izquierdo, píxel de un color diferente al píxel de la coordenada @initCoord)
            startPoint = new System.Windows.Point(initCoord.X, yPos);
            double xPos = searchLeftBorder(bmImage, startPoint/*startPoint*//*, c*/);
            // Tercer desplazamiento. 
            // V1.0: Comienza en el mismo punto que el primer desplazamiento: @initCoord.
            // V2.0: Siguiendo el contorno
            //startPoint = new System.Windows.Point(initCoord.X, yPos);
            double bottomBorderPos = searchBottomBorder(bmImage, initCoord);
            // Cuarto desplazamiento. Comienzo: @startPoint.
            startPoint = new System.Windows.Point(initCoord.X, bottomBorderPos);
            double rightBorderPos = searchRightBorder(bmImage, startPoint);

            FakeTextBox.Margin = new Thickness(xPos, yPos, 0, 0);
            FakeTextBox.Width = rightBorderPos - xPos;
            FakeTextBox.Height = bottomBorderPos - yPos;

            float br = GetBrightness(c);

            //White background: 0.9862745
            //Blue background : 0.4
            if (br > 0.8)
            {
                foreColor = Colors.Black;
            }
            else
            {
                foreColor = Colors.White;
            }

            FakeTextBox.Background = new SolidColorBrush(c);
            FakeTextBox.Foreground = new SolidColorBrush(foreColor);
        }
        private double searchTopBorder(BitmapImage bmImage, System.Windows.Point currentCoord/*, System.Windows.Media.Color c*/)
        {
            double newYPos = currentCoord.Y;
            if ((newYPos - 1) > 0)
            {
                newYPos--;

                System.Windows.Point nextCoord = new System.Windows.Point(currentCoord.X, newYPos);
                System.Windows.Media.Color previousColor = GetCoordinateColor(/*bmImage, */ currentCoord.X, currentCoord.Y),
                                                newCoordColor = GetCoordinateColor(/*bmImage, */ nextCoord.X, nextCoord.Y);

                float totalDiff = GetColorDifference(previousColor, newCoordColor);

                if (totalDiff < diffThreshold /*|| (newCoordColor.R == 0 && newCoordColor.G == 0 && newCoordColor.B == 0)*/)
                    newYPos = searchTopBorder(bmImage, nextCoord/*, c*/);
            }

            return newYPos;
        }
        private double searchLeftBorder(BitmapImage bmImage, System.Windows.Point currentCoord/*, System.Windows.Media.Color c*/)
        {
            double newXPos = currentCoord.X;
            if ((newXPos - 1) > 0)
            {
                newXPos--;

                System.Windows.Point nextCoord = new System.Windows.Point(newXPos, currentCoord.Y);
                System.Windows.Media.Color previousColor = GetCoordinateColor(/*bmImage, */currentCoord.X, currentCoord.Y), 
                                                newCoordColor = GetCoordinateColor(/*bmImage, */nextCoord.X, nextCoord.Y);
                float totalDiff = GetColorDifference(previousColor, newCoordColor);

                if (newXPos < 390)
                    Console.WriteLine("X coord = (" + newXPos + "). Current diff = (" + totalDiff + ")");

                if (totalDiff < diffThreshold /*|| (newCoordColor.R == 0 && newCoordColor.G == 0 && newCoordColor.B == 0)*/)
                    newXPos = searchLeftBorder(bmImage, nextCoord/*, c*/);
            }

            return newXPos;
        }
        private double searchBottomBorder(BitmapImage bmImage, System.Windows.Point currentCoord/*, System.Windows.Media.Color c*/)
        {
            double bottomPos = currentCoord.Y;
            if ((bottomPos + 1) < bmImage.PixelHeight)
            {
                bottomPos++;

                System.Windows.Point nextCoord = new System.Windows.Point(currentCoord.X, bottomPos);
                System.Windows.Media.Color previousColor = GetCoordinateColor(/*bmImage, */currentCoord.X, currentCoord.Y),
                                                newCoordColor = GetCoordinateColor(/*bmImage, */nextCoord.X, nextCoord.Y);

                float totalDiff = GetColorDifference(previousColor, newCoordColor);

                if (totalDiff < diffThreshold  /*|| (newCoordColor.R == 0 && newCoordColor.G == 0 && newCoordColor.B == 0)*/)
                    bottomPos = searchBottomBorder(bmImage, nextCoord/*, c*/);
            }

            return bottomPos;
        }
        private double searchRightBorder(BitmapImage bmImage, System.Windows.Point currentCoord/*, System.Windows.Media.Color c*/)
        {
            double rightBorderPos= currentCoord.X;
            if ((rightBorderPos + 1) < bmImage.PixelWidth)
            {
                rightBorderPos++;

                System.Windows.Point nextCoord = new System.Windows.Point(rightBorderPos, currentCoord.Y);
                System.Windows.Media.Color previousColor = GetCoordinateColor(/*bmImage, */currentCoord.X, currentCoord.Y),
                                                newCoordColor = GetCoordinateColor(/*bmImage, */nextCoord.X, nextCoord.Y);

                float totalDiff = GetColorDifference(previousColor, newCoordColor);

                if (totalDiff < diffThreshold)
                    rightBorderPos = searchRightBorder(bmImage, nextCoord/*, c*/);
            }

            return rightBorderPos;
        }
        private float GetColorDifference(System.Windows.Media.Color c1, System.Windows.Media.Color c2)
        {
            int diffRed = Math.Abs(c1.R - c2.R);
            int diffGreen = Math.Abs(c1.G - c2.G);
            int diffBlue = Math.Abs(c1.B - c2.B);
            float pctDiffRed = (float)diffRed / 255;
            float pctDiffGreen = (float)diffGreen / 255;
            float pctDiffBlue = (float)diffBlue / 255;

            return (pctDiffRed + pctDiffGreen + pctDiffBlue) / 3 /* (* 100)*/;
        }
        private System.Windows.Media.Color GetCoordinateColor(/*BitmapImage bmImage, */double iX, double iY)
        {
            System.Windows.Media.Color newColor;

            //using (MemoryStream outStream = new MemoryStream())
            //{
            //    BitmapEncoder enc = new BmpBitmapEncoder();
            //    enc.Frames.Add(BitmapFrame.Create(bmImage));
            //    enc.Save(outStream);
            //    System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(outStream);

                //Bitmap b = new Bitmap(bitmap);
                System.Drawing.Color color = _Image.GetPixel((int)iX, (int)iY);

                //bitmap.Dispose();
                //b.Dispose();

                newColor = System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B);
            //}


            return newColor;
        }
        private Bitmap BitmapImage2Bitmap(BitmapImage bitmapImage)
        {
            // BitmapImage bitmapImage = new BitmapImage(new Uri("../Images/test.png", UriKind.Relative));

            using (MemoryStream outStream = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapImage));
                enc.Save(outStream);
                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(outStream);

                return new Bitmap(bitmap);
            }
        }
    }
}