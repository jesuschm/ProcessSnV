using RPAValidator.Model;
using RPAValidator.Views;
using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace RPAValidator
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        FakeUI fakeUI;
        String logPath = @"D:\Personal\TFG\RPAValidator\RPAValidator\Resources\log.csv";

        public MainWindow()
        {
            InitializeComponent();
            Build();
        }

        public void Build()
        {
            BtnPlay.IsEnabled = false;
            fakeUI = new FakeUI(this);
        }

        private void BtnPlay_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            fakeUI.Start();
        }

        private void BtnLoadUI_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(logPath))
            {
                using (var reader = new StreamReader(logPath))
                {
                    string line = reader.ReadLine();
                    string[] readValues;
                    List<String> values = new List<String>();

                    int expectedEventsCount = 0;

                    String keystrokeImage = "";
                    while (!reader.EndOfStream)
                    {
                        line = reader.ReadLine();
                        readValues = line.Split(';');
                        // [2] App
                        if (readValues[2] != "" && !readValues[2].Contains("Aquiles") && (readValues[10].Contains("jpg") || keystrokeImage != ""))
                        {
                            values.Clear();
                            String eventType = readValues[6], eInfo = readValues[8], xCoord = "0", yCoord = "0";

                            if (eventType.Equals(Event.EventType.Cursor.ToString()))
                            {
                                if (keystrokeImage != "") { 
                                    // Tras asociarle la imagen del keystroke al evento click sin imagen, se reinicia la variable que contiene dicha imagen.
                                    readValues[10] = keystrokeImage;
                                    keystrokeImage = "";
                                }

                                xCoord = eInfo.Substring(eInfo.IndexOf("}"), (eInfo.Length - eInfo.IndexOf("-"))).Replace(" ", "").Replace("}", "").Replace("-", " ");
                                yCoord = eInfo.Substring(eInfo.IndexOf("-")).Replace(" ", "").Replace("}", "").Replace("-", " ").Replace(" ", "");
                            }
                            else
                            {
                                // El evento click de después de un keystroke, no contiene imagen asociada. Se le asociará la imagen del propio keystroke.
                                keystrokeImage = readValues[10];
                                eInfo = eInfo.Replace("{", "").Replace("}", "").Replace("SHIFT", "").Replace("ENTER", "").Replace("CONTROL", "");
                            }

                            values.Add(readValues[0]); // [0] Id
                            //values.Add(readValues[3]); // [3] X coord
                            //values.Add(readValues[4]); // [4] Y coord
                            values.Add(xCoord);
                            values.Add(yCoord);
                            values.Add(eventType); // [6] Event type
                            values.Add(eInfo); // [8] Event type information
                            values.Add(readValues[10]); // [10] Image name

                            fakeUI.AddExpectedEvent(values);
                            expectedEventsCount++;
                        }
                    }

                    LbLoadMessage.Content = "Se han añadido " + expectedEventsCount + " eventos esperados.";
                    BtnPlay.IsEnabled = true;
                }
            }
        }
        public void TestFinished(bool interrupted)
        {
            LbLoadMessage.Content = "Test finished";
            if (interrupted)
                LbLoadMessage.Content = "Robot fails";

            fakeUI.Hide();
            fakeUI.Restart();

            this.Show();
        }
        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
