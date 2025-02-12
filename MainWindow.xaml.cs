using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Globalization;
using System.Transactions;
using System.Collections.Generic;
using System.Text;
using Azure;

using System.Diagnostics;

namespace ComtradeFileReasAndInsert
{
    public partial class MainWindow : Window
    {
        private string filePath;
        private string fileName;
        private string fileExtension;

        public struct AnalogData
        {
            public int ChannelIndexNumber;
            public string ChannelId;
            public string PhaseId;
            public string Ccbm;
            public string ChannelUnits;
            public double ChannelMultiplier;
            public double ChannelOffset;
            public double ChannelSkew;
            public double MinimumLimit;
            public double MaximumLimit;
            public double ChannelRatioPrimary;
            public double ChannelRatioSecondary;
            public string DataPrimarySecondary;
        }

        public struct DigitalData
        {
            public int ChannelNumber;
            public string ChannelId;
            public string PhaseId;
            public string Ccbm;
            public int NormalState;
        }

        public struct ComtradeData
        {
            public string Station;
            public string DeviceId;
            public int CfgVersion;
            public float LineFrequency;
            public int SampleRateCount;
            public float SampleRate;
            public int LastSampleRate;
            public DateTime FirstTimeStamp;

            public DateTime TriggerTimeStamp;

            public string DataType;
            public double TimeMultiplier;
            public string TimeCode;
            public string LocalCode;
            public string TimeQualityIndicatorCode;
            public int LeapSecondIndicator;
        }

        public struct ComtradeData1
        {
            public int TotalSignalCount;
            public int AnalogSignalCount;
            public int DigitalSignalCount;
        }

        // Global Variables
        public static ComtradeData Comtrade;
        public static List<AnalogData> Analog = new List<AnalogData>();
        public static List<DigitalData> Digital = new List<DigitalData>();
        public static ComtradeData1 Comtrade1;
        public static List<string> Words = new List<string>();
        public static int ComtradeIndex;
        private static List<int> valueArray = new List<int>();
        public static List<int> timeArray = new List<int>();
        public static List<int> datIndexArray = new List<int>();
        public int[] intValues;


        public MainWindow()
        {
            InitializeComponent();
        }

        // Method to extract revised year from a line 2013 cfg
        public static void ExtractRevisedYear(string line)
        {
            var tokens = line.Split(',');

            if (tokens[0].Length < 1 || tokens[1].Length < 1 || tokens[2].Length < 1)
            {
                MessageBox.Show("Error: Invalid file format. Extract Revised Year");
                Application.Current.Shutdown();
            }
            if (tokens[0].Length > 64 || tokens[1].Length > 64)
            {
                MessageBox.Show("Error: Invalid file format.more than expected characters");
                Application.Current.Shutdown();
            }
            // Check if there are at least 3 tokens
            if (tokens.Length > 2)
            {
                if (int.TryParse(tokens[2].Trim(), out int revisionYear))
                {
                    Comtrade.CfgVersion = revisionYear;
                }
                Comtrade.Station = tokens[0];
                Comtrade.DeviceId = tokens[1];
                if (Comtrade.CfgVersion != 1999 && Comtrade.CfgVersion != 2013 && Comtrade.CfgVersion != 1991)
                {
                    MessageBox.Show("Error: Invalid file format. Comtrade Version");
                    Application.Current.Shutdown();
                }
            }
            else
            {
                MessageBox.Show("Error: Invalid file format. Extract Revised Year");
                Application.Current.Shutdown();

            }

        }

        //Extracting first line of 1999 cfg
        public static void ExtractRevisedYear1999(string line)
        {
            var tokens = line.Split(',');


            if (tokens[0].Length > 64 || tokens[1].Length > 64)
            {
                MessageBox.Show("Error: Invalid file format.more than expected characters");
                Application.Current.Shutdown();
            }
            // Check if there are at least 3 tokens
            if (tokens.Length > 2)
            {
                if (int.TryParse(tokens[2].Trim(), out int revisionYear))
                {
                    Comtrade.CfgVersion = revisionYear;
                }
                Comtrade.Station = tokens[0];
                Comtrade.DeviceId = tokens[1];

            }
            if (Comtrade.CfgVersion != 1999)
            {
                MessageBox.Show("Error:Invalid file"); Application.Current.Shutdown();
            }

        }



        // Method to parse a sentence into words 2013
        public static void ComtradeParse(string sentence)
        {
            string[] tokens = sentence.Split(',');


            foreach (var token in tokens)
            {
                if (Words.Count >= 14) return;
                Words.Add(token);

            }
        }

        // Method to parse a sentence into words 1999
        public static void ComtradeParse1999(string sentence)
        {
            string[] tokens = sentence.Split(',');


            foreach (var token in tokens)
            {
                if (Words.Count >= 10) return;
                Words.Add(token);

            }
        }


        // Process parsed words into ComtradeData 1999
        public static void ProcessWords1999(List<string> words)
        {
            if (words.Count < 6 || words.Count > 10)
            {
                MessageBox.Show("Error: Invalid file format. Comtrade Data");
                Application.Current.Shutdown();
            }

            if (string.IsNullOrEmpty(words[1])) { MessageBox.Show("Error: SampleRateCount is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[2])) { MessageBox.Show("Error: Samplerate is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[3])) { MessageBox.Show("Error: Last Sample rate is empty"); Application.Current.Shutdown(); }

            if (string.IsNullOrEmpty(words[8])) { MessageBox.Show("Error:Data type is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[9])) { MessageBox.Show("Error: Time multiplier is empty"); Application.Current.Shutdown(); }

            if (int.Parse(words[0]) < 0) { MessageBox.Show("Error: Line frequency is negative"); Application.Current.Shutdown(); }
            Comtrade.LineFrequency = float.Parse(words[0]);
            Comtrade.SampleRateCount = int.Parse(words[1]);
            Comtrade.SampleRate = float.Parse(words[2]);
            Comtrade.LastSampleRate = int.Parse(words[3]);
            string time1 = words[4] + " " + words[5];
            Comtrade.FirstTimeStamp = DateTime.Parse(time1);

            string time2 = words[6] + " " + words[7];

            Comtrade.TriggerTimeStamp = DateTime.Parse(time2);
            if (Comtrade.TriggerTimeStamp < Comtrade.FirstTimeStamp)
            {
                MessageBox.Show("Error: Trigger Time Stamp is less than First sample time");
                Application.Current.Shutdown();
            }

            Comtrade.DataType = words[8];

            if (Comtrade.DataType != "ASCII" && Comtrade.DataType != "BINARY" && Comtrade.DataType != "BINARY32" && Comtrade.DataType != "FLOAT32")
            {
                MessageBox.Show("Error: Invalid file format.Data Type");
                Application.Current.Shutdown();
            }
            Comtrade.TimeMultiplier = double.Parse(words[9]);

        }



        // Process parsed words into ComtradeData 2013
        public static void ProcessWords(List<string> words)
        {
            if (words.Count != 14)
            {
                MessageBox.Show("Error: Invalid file format. Comtrade Data");
                Application.Current.Shutdown();
            }
            if (string.IsNullOrEmpty(words[0])) { MessageBox.Show("Error: Line frequency is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[1])) { MessageBox.Show("Error: SampleRateCount is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[2])) { MessageBox.Show("Error: Samplerate is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[3])) { MessageBox.Show("Error: Last Sample rate is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[4])) { MessageBox.Show("Error: First time stamp is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[6])) { MessageBox.Show("Error: Last time stamp is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[8])) { MessageBox.Show("Error:Data type is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[9])) { MessageBox.Show("Error: Time multiplier is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[10])) { MessageBox.Show("Error:Time code is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[11])) { MessageBox.Show("Error: Local code is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[12])) { MessageBox.Show("Error:TimeQualityIndicatorCode is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(words[13])) { MessageBox.Show("Error: LeapSecondIndicator is empty"); Application.Current.Shutdown(); }
            if (int.Parse(words[0]) < 0) { MessageBox.Show("Error: Line frequency is negative"); Application.Current.Shutdown(); }
            Comtrade.LineFrequency = float.Parse(words[0]);
            Comtrade.SampleRateCount = int.Parse(words[1]);
            Comtrade.SampleRate = float.Parse(words[2]);
            Comtrade.LastSampleRate = int.Parse(words[3]);
            string time1 = words[4] + " " + words[5];
            Comtrade.FirstTimeStamp = DateTime.Parse(time1);

            string time2 = words[6] + " " + words[7];

            Comtrade.TriggerTimeStamp = DateTime.Parse(time2);
            if (Comtrade.TriggerTimeStamp < Comtrade.FirstTimeStamp)
            {
                MessageBox.Show("Error: Trigger Time Stamp is less than First sample time");
                Application.Current.Shutdown();
            }

            Comtrade.DataType = words[8];
            MessageBox.Show(Comtrade.DataType);
            if (Comtrade.DataType != "ASCII" && Comtrade.DataType != "BINARY" && Comtrade.DataType != "BINARY32" && Comtrade.DataType != "FLOAT32")
            {
                MessageBox.Show("Error: Invalid file format.Data Type");
                Application.Current.Shutdown();
            }
            Comtrade.TimeMultiplier = double.Parse(words[9]);
            Comtrade.TimeCode = words[10];
            Comtrade.LocalCode = words[11];
            Comtrade.TimeQualityIndicatorCode = words[12];
            Comtrade.LeapSecondIndicator = int.Parse(words[13]);
        }

        // Count signals for both 2013 & 1999
        public static void SignalCounting(string sentence)
        {
            var tokens = sentence.Split(',');
            if (tokens.Length != 3)
            {
                MessageBox.Show("Error: Invalid file format.Signal Count");
                Application.Current.Shutdown();

            }
            Comtrade1.TotalSignalCount = int.Parse(tokens[0]);
            if (Comtrade1.TotalSignalCount > 999999 || Comtrade1.TotalSignalCount < 1)
            {
                MessageBox.Show("Error:Channel number should be less than 999999.");
                Application.Current.Shutdown();
            }
            if (tokens[1].Length < 2 || tokens[2].Length < 2)
            {
                MessageBox.Show("Error:Analog channel and Digital number should be atleast 2 characters.");
                Application.Current.Shutdown();
            }

            if (tokens[1][tokens[1].Length - 1] != 'A' || tokens[2][tokens[2].Length - 1] != 'D')
            {
                MessageBox.Show("Error:Invalid Character"); Application.Current.Shutdown();
            }
            if (tokens[1].Length > 2)
            {
                Comtrade1.AnalogSignalCount = int.Parse(tokens[1].Substring(0, 2));
            }
            else
            {
                Comtrade1.AnalogSignalCount = int.Parse(tokens[1].Substring(0, 1));
            }
            if (tokens[2].Length > 2)
            {
                Comtrade1.DigitalSignalCount = int.Parse(tokens[2].Substring(0, 2));
            }
            else
            {
                Comtrade1.DigitalSignalCount = int.Parse(tokens[2].Substring(0, 1));

            }
            if (Comtrade1.TotalSignalCount != (Comtrade1.AnalogSignalCount + Comtrade1.DigitalSignalCount))
            {
                MessageBox.Show("Error: Invalid file format.Total signals");
                Application.Current.Shutdown();
            }

        }



        //for 2013 CFG
        public static void ParseAndStoreAnalogData(string line)
        {
            var tokens = line.Split(',');
            if (tokens.Length != 13)
            {
                MessageBox.Show("Error: Invalid file format. Analog Signal");
                Application.Current.Shutdown();
            }
            if (string.IsNullOrEmpty(tokens[0])) { MessageBox.Show("Error: Channel Index Number is empty"); Application.Current.Shutdown(); }
            if (tokens[0].Length > 6)
            {
                MessageBox.Show("Error"); Application.Current.Shutdown();
            }
            if (!int.TryParse(tokens[0], out int result)) { MessageBox.Show("Error: Channel Index Number is not in correct format"); Application.Current.Shutdown(); }

            if (tokens[1].Length > 128) { MessageBox.Show("Error"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[1])) { MessageBox.Show("Error: Channel Identifier is empty"); Application.Current.Shutdown(); }

            if (string.IsNullOrEmpty(tokens[4])) { MessageBox.Show("Error: Channel Units  is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[5])) { MessageBox.Show("Error: Channel Multiplier is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[6])) { MessageBox.Show("Error: Channel offset is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[7])) { MessageBox.Show("Error: Skew is empty"); Application.Current.Shutdown(); }
            if (float.Parse(tokens[8]) > float.Parse(tokens[9])) { MessageBox.Show("Error: Min is greater than max"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[12])) { MessageBox.Show("Error: Channel type is empty"); Application.Current.Shutdown(); }
            if (tokens[12] != "s" && tokens[12] != "p" && tokens[12] != "P" && tokens[12] != "S") { MessageBox.Show("Error: Channel type is not in correct format"); Application.Current.Shutdown(); }

            int channelIndexNumber = int.Parse(tokens[0]);

            // Check for duplication or non-sequential ChannelIndexNumber
            if (Analog.Any(a => a.ChannelIndexNumber == channelIndexNumber))
            {
                MessageBox.Show("Error: Duplicate Channel Index Number");
                Application.Current.Shutdown();
            }
            if (Analog.Count > 0 && channelIndexNumber != Analog.Last().ChannelIndexNumber + 1)
            {
                MessageBox.Show("Error: Non-sequential Channel Index Number");
                Application.Current.Shutdown();
            }

            AnalogData analog = new AnalogData
            {
                ChannelIndexNumber = channelIndexNumber,
                ChannelId = tokens[1],
                PhaseId = tokens[2],
                Ccbm = tokens[3],
                ChannelUnits = tokens[4],
                ChannelMultiplier = double.Parse(tokens[5]),
                ChannelOffset = double.Parse(tokens[6]),
                ChannelSkew = double.Parse(tokens[7]),
                MinimumLimit = double.Parse(tokens[8]),
                MaximumLimit = double.Parse(tokens[9]),
                ChannelRatioPrimary = double.Parse(tokens[10]),
                ChannelRatioSecondary = double.Parse(tokens[11]),
                DataPrimarySecondary = tokens[12]
            };
            Analog.Add(analog);
        }


        //for 1999 cfg

        public static void ParseAndStoreAnalogData1999(string line)
        {
            var tokens = line.Split(',');
            if (tokens.Length != 13)
            {
                MessageBox.Show("Error: Invalid file format. Analog Signal");
                Application.Current.Shutdown();
            }
            if (string.IsNullOrEmpty(tokens[0])) { MessageBox.Show("Error: Channel Index Number is empty"); Application.Current.Shutdown(); }
            if (tokens[0].Length > 6)
            {
                MessageBox.Show("Error"); Application.Current.Shutdown();
            }
            if (!int.TryParse(tokens[0], out int result)) { MessageBox.Show("Error: Channel Index Number is not in correct format"); Application.Current.Shutdown(); }

            if (tokens[1].Length > 128) { MessageBox.Show("Error"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[1])) { MessageBox.Show("Error: Channel Identifier is empty"); Application.Current.Shutdown(); }

            if (string.IsNullOrEmpty(tokens[4])) { MessageBox.Show("Error: Channel Units  is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[5])) { MessageBox.Show("Error: Channel Multiplier is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[6])) { MessageBox.Show("Error: Channel offset is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[8])) { MessageBox.Show("Error: Invalid file format"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[9])) { MessageBox.Show("Error: Invalid file format"); Application.Current.Shutdown(); }
            //  if (float.Parse(tokens[8]) > float.Parse(tokens[9])) { MessageBox.Show("Error: Min is greater than max"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[12])) { MessageBox.Show("Error: Channel type is empty"); Application.Current.Shutdown(); }
            if (tokens[12] != "s" && tokens[12] != "p" && tokens[12] != "P" && tokens[12] != "S") { MessageBox.Show("Error: Channel type is not in correct format"); Application.Current.Shutdown(); }

            int channelIndexNumber = int.Parse(tokens[0]);

            // Check for duplication or non-sequential ChannelIndexNumber
            if (Analog.Any(a => a.ChannelIndexNumber == channelIndexNumber))
            {
                MessageBox.Show("Error: Duplicate Channel Index Number");
                Application.Current.Shutdown();
            }
            if (Analog.Count > 0 && channelIndexNumber != Analog.Last().ChannelIndexNumber + 1)
            {
                MessageBox.Show("Error: Non-sequential Channel Index Number");
                Application.Current.Shutdown();
            }

            AnalogData analog = new AnalogData
            {
                ChannelIndexNumber = channelIndexNumber,
                ChannelId = tokens[1],
                PhaseId = tokens[2],
                Ccbm = tokens[3],
                ChannelUnits = tokens[4],
                ChannelMultiplier = double.Parse(tokens[5]),
                ChannelOffset = double.Parse(tokens[6]),
                ChannelSkew = double.Parse(tokens[7]),
                MinimumLimit = double.Parse(tokens[8]),
                MaximumLimit = double.Parse(tokens[9]),
                ChannelRatioPrimary = double.Parse(tokens[10]),
                ChannelRatioSecondary = double.Parse(tokens[11]),
                DataPrimarySecondary = tokens[12]
            };
            Analog.Add(analog);
        }






        public static void ParseAndStoreDigitalData(string line)
        {
            var tokens = line.Split(',');

            // Check if there are enough tokens before processing
            if (tokens.Length != 5)
            {
                MessageBox.Show("Error: Invalid file format. Digital Signal");
                Application.Current.Shutdown();
            }

            int channelIndexNumber = int.Parse(tokens[0]);

            // Check for duplication or non-sequential ChannelIndexNumber
            if (Digital.Any(a => a.ChannelNumber == channelIndexNumber))
            {
                MessageBox.Show("Error: Duplicate Channel Index Number");
                Application.Current.Shutdown();
            }
            if (Digital.Count > 0 && channelIndexNumber != Digital.Last().ChannelNumber + 1)
            {
                MessageBox.Show("Error: Non-sequential Channel Index Number");
                Application.Current.Shutdown();
            }

            if (string.IsNullOrEmpty(tokens[0])) { MessageBox.Show("Error: Channel Number is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[1])) { MessageBox.Show("Error: Channel Identifier is empty"); Application.Current.Shutdown(); }
            if (string.IsNullOrEmpty(tokens[4])) { MessageBox.Show("Error:Normal state is empty"); Application.Current.Shutdown(); }
            if (tokens[4] != "0" && tokens[4] != "1") { MessageBox.Show("Error: Normal State is not int correct formate"); Application.Current.Shutdown(); }
            DigitalData digital = new DigitalData
            {
                ChannelNumber = int.Parse(tokens[0]),
                ChannelId = tokens[1],
                PhaseId = tokens[2],
                Ccbm = tokens[3],
                NormalState = int.Parse(tokens[4])
            };
            Digital.Add(digital);

        }

        //for 1999 cfg
        public static void ParseAndStoreDigitalData1999(string line)
        {
            var tokens = line.Split(',');

            // Check if there are enough tokens before processing
            if (tokens.Length != 5)
            {
                MessageBox.Show("Error: Invalid file format. Digital Signal");
                Application.Current.Shutdown();
            }
            if (!int.TryParse(tokens[0], out int result)) { MessageBox.Show("Error: Channel Index Number is not in correct format"); Application.Current.Shutdown(); }

            int channelIndexNumber = int.Parse(tokens[0]);

            // Check for duplication or non-sequential ChannelIndexNumber
            if (Digital.Any(a => a.ChannelNumber == channelIndexNumber))
            {
                MessageBox.Show("Error: Duplicate Channel Index Number");
                Application.Current.Shutdown();
            }
            if (Digital.Count > 0 && channelIndexNumber != Digital.Last().ChannelNumber + 1)
            {
                MessageBox.Show("Error: Non-sequential Channel Index Number");
                Application.Current.Shutdown();
            }

            if (string.IsNullOrEmpty(tokens[0])) { MessageBox.Show("Error: Channel Number is empty"); Application.Current.Shutdown(); }

            if (string.IsNullOrEmpty(tokens[4])) { MessageBox.Show("Error:Normal state is empty"); Application.Current.Shutdown(); }
            if (tokens[4] != "0" && tokens[4] != "1") { MessageBox.Show("Error: Normal State is not int correct format"); Application.Current.Shutdown(); }
            DigitalData digital = new DigitalData
            {
                ChannelNumber = int.Parse(tokens[0]),
                ChannelId = tokens[1],
                PhaseId = tokens[2],
                Ccbm = tokens[3],
                NormalState = int.Parse(tokens[4])
            };
            Digital.Add(digital);

        }




        public static void AsciiDat(string line, int Analogcount, int DigitalCount)
        {
            string connectionString = "Data Source=SANAL-PROEDISON\\SQLEXPRESS;Initial Catalog=Demo;User ID=sa;Password=mypassword;Encrypt=False;";
            int k = 0;
            string[] values = line.Split(',');
            int[] intValues = Array.ConvertAll(values, s =>
            {
                if (values[0].Length == 0)
                {
                    MessageBox.Show("Error: DatIndex is empty");
                    Application.Current.Shutdown();
                }
                return int.Parse(s);
            });

            if (datIndexArray.Count == 0)
            {
                datIndexArray.Add(intValues[0]);
            }
            else if (datIndexArray[datIndexArray.Count - 1] + 1 != intValues[0])
            {
                MessageBox.Show("Error: DatIndex must be sequential");
                Application.Current.Shutdown();
            }
            else
            {
                datIndexArray.Add(intValues[0]);
            }

            if (timeArray.Count == 0)
            {
                timeArray.Add(intValues[1]);
            }
            else if (timeArray[timeArray.Count - 1] > intValues[1])
            {
                MessageBox.Show("Error: Time value must be greater than the previous values.");
                Application.Current.Shutdown();
            }
            else
            {
                timeArray.Add(intValues[1]);
            }

            string query4 = "INSERT INTO AnalogDat(ComtradeIndex,AnalogIndex,DatIndex,Time,Value,Result) VALUES (@ComtradeIndex,@AnalogIndex,@DatIndex,@Time,@Value,@Result)";
            string query5 = "INSERT INTO DigitalDat(ComtradeIndex,DigitalIndex,DatIndex,Time,Value) VALUES (@ComtradeIndex,@DigitalIndex,@DatIndex,@Time,@Value)";
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlTransaction transaction = con.BeginTransaction())
                {
                    for (int i = 2; i < Analogcount; i++)
                    {
                        double result = 0;
                        if (Analog[i - 2].DataPrimarySecondary.Equals("S"))
                        {
                            result = ((intValues[i] * Analog[i - 2].ChannelMultiplier) + Analog[i - 2].ChannelOffset) * Analog[i - 2].ChannelRatioPrimary;
                        }
                        else
                        {
                            result = (intValues[i] * Analog[i - 2].ChannelMultiplier) + Analog[i - 2].ChannelOffset;
                        }
                        using (SqlCommand cmdAnalogDat = new SqlCommand(query4, con, transaction))
                        {
                            cmdAnalogDat.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                            cmdAnalogDat.Parameters.AddWithValue("@AnalogIndex", Analog[i - 2].ChannelIndexNumber);
                            cmdAnalogDat.Parameters.AddWithValue("@DatIndex", intValues[0]);
                            cmdAnalogDat.Parameters.AddWithValue("@Time", intValues[1]);
                            cmdAnalogDat.Parameters.AddWithValue("@Value", intValues[i]);
                            cmdAnalogDat.Parameters.AddWithValue("@Result", result);

                            cmdAnalogDat.ExecuteNonQuery();
                        }
                    }

                    // Uncomment and update the Digital data insertion logic if needed
                    for (int i = 2 + Comtrade1.AnalogSignalCount; i < intValues.Length; i++)
                    {
                        using (SqlCommand cmdDigitalDat = new SqlCommand(query5, con, transaction))
                        {
                            cmdDigitalDat.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                            cmdDigitalDat.Parameters.AddWithValue("@DigitalIndex", Digital[k].ChannelNumber);
                            cmdDigitalDat.Parameters.AddWithValue("@DatIndex", intValues[0]);
                            cmdDigitalDat.Parameters.AddWithValue("@Time", intValues[1]);
                            cmdDigitalDat.Parameters.AddWithValue("@Value", intValues[i]);

                            cmdDigitalDat.ExecuteNonQuery();
                        }
                        k++;
                    }
                    transaction.Commit();
                }
            }
        }
        public static void BinaryDat(string[] hexChunk, int AnalogCount, int DigitalCount)
        {
            string connectionString = "Data Source=SANAL-PROEDISON\\SQLEXPRESS;Initial Catalog=Demo;User ID=sa;Password=mypassword;Encrypt=False;";
            string query4 = "INSERT INTO AnalogDat(ComtradeIndex,AnalogIndex,DatIndex,Time,Value,Result) VALUES (@ComtradeIndex,@AnalogIndex,@DatIndex,@Time,@Value,@Result)";
            string query5 = "INSERT INTO DigitalDat(ComtradeIndex,DigitalIndex,DatIndex,Time,Value) VALUES (@ComtradeIndex,@DigitalIndex,@DatIndex,@Time,@Value)";
            string indexHex = hexChunk[3] + hexChunk[2] + hexChunk[1] + hexChunk[0];
            if (indexHex.Length == 0)
            {
                MessageBox.Show("Error: Invalid file format. IndexHex");
                Application.Current.Shutdown();
            }
            string timeStampHex = hexChunk[7] + hexChunk[6] + hexChunk[5] + hexChunk[4];

            int num = Convert.ToInt32(indexHex, 16);
            int num1 = Convert.ToInt32(timeStampHex, 16);

            if (datIndexArray.Count == 0)
            {
                datIndexArray.Add(num);
            }
            else if (datIndexArray[datIndexArray.Count - 1] + 1 != num)
            {
                MessageBox.Show("Error: DatIndex must be sequential");
                Application.Current.Shutdown();
            }
            else
            {
                datIndexArray.Add(num);
            }

            if (timeArray.Count == 0)
            {
                timeArray.Add(num1);
            }
            else if (timeArray[timeArray.Count - 1] > num1)
            {
                MessageBox.Show("Error: Time value must be greater than the previous values.");
                Application.Current.Shutdown();
            }
            else
            {
                timeArray.Add(num1);
            }

            int k = 0;
            int d = 0;
            int ascii;

            for (int i = 8; i < 8 + (2 * AnalogCount); i += 2)
            {
                string x = hexChunk[i + 1] + hexChunk[i];
                int decimalValue = Convert.ToInt32(x, 16);

                string binaryValue = Convert.ToString(decimalValue, 2).PadLeft(16, '0');
                string invertedBits = "";
                for (int j = 0; j < binaryValue.Length; j++)
                {
                    invertedBits += binaryValue[j] == '1' ? '0' : '1'; // Finding One's Complement
                }
                int a = Convert.ToInt32(invertedBits, 2);
                int ans = (a + 1) * (-1);

                if (binaryValue[0] == '1')
                {
                    ascii = ans;
                }
                else
                {
                    ascii = decimalValue;
                }

                double result = 0;
                if (Analog[k].DataPrimarySecondary.Equals("S"))
                {
                    result = ((ascii * Analog[k].ChannelMultiplier) + Analog[k].ChannelOffset) * Analog[k].ChannelRatioPrimary;
                }
                else
                {
                    result = (ascii * Analog[k].ChannelMultiplier) + Analog[k].ChannelOffset;
                }

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    using (SqlTransaction transaction = con.BeginTransaction())
                    {
                        using (SqlCommand cmdAnalogDat = new SqlCommand(query4, con, transaction))
                        {
                            cmdAnalogDat.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                            cmdAnalogDat.Parameters.AddWithValue("@AnalogIndex", Analog[k].ChannelIndexNumber);
                            cmdAnalogDat.Parameters.AddWithValue("@DatIndex", num);
                            cmdAnalogDat.Parameters.AddWithValue("@Time", num1);
                            cmdAnalogDat.Parameters.AddWithValue("@Value", ascii);
                            cmdAnalogDat.Parameters.AddWithValue("@Result", result);

                            cmdAnalogDat.ExecuteNonQuery();
                        }

                        transaction.Commit();
                        con.Close();
                    }
                }
                k++;
            }
            for (int i = 8 + (2 * AnalogCount); i < hexChunk.Length; i = i + 2)
            {

                string statusBitsHex = hexChunk[i + 1] + hexChunk[i];
                int bits = Convert.ToInt32(statusBitsHex, 16);
                string statusBitsBin = Convert.ToString(bits, 2).PadLeft(16, '0');

                for (int j = statusBitsBin.Length - 1; j >= 0; j--)
                {

                    using (SqlConnection con = new SqlConnection(connectionString))
                    {
                        con.Open();
                        using (SqlTransaction transaction = con.BeginTransaction())
                        {
                            using (SqlCommand cmdDigitalDat = new SqlCommand(query5, con, transaction))
                            {
                                cmdDigitalDat.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                                cmdDigitalDat.Parameters.AddWithValue("@DigitalIndex", Digital[d].ChannelNumber);
                                cmdDigitalDat.Parameters.AddWithValue("@DatIndex", num);
                                cmdDigitalDat.Parameters.AddWithValue("@Time", num1);
                                cmdDigitalDat.Parameters.AddWithValue("@Value", statusBitsBin[j]);
                                cmdDigitalDat.ExecuteNonQuery();
                            }
                            transaction.Commit();
                            con.Close();
                        }
                    }
                    d++;

                }
            }


        }
        public static void Binary32Dat(string[] hexChunk, int AnalogCount, int DigitalCount)
        {

            string connectionString = "Data Source=SANAL-PROEDISON\\SQLEXPRESS;Initial Catalog=Demo;User ID=sa;Password=mypassword;Encrypt=False;";
            string query4 = "INSERT INTO AnalogDat(ComtradeIndex,AnalogIndex,DatIndex,Time,Value,Result) VALUES (@ComtradeIndex,@AnalogIndex,@DatIndex,@Time,@Value,@Result)";
            string query5 = "INSERT INTO DigitalDat(ComtradeIndex,DigitalIndex,DatIndex,Time,Value) VALUES (@ComtradeIndex,@DigitalIndex,@DatIndex,@Time,@Value)";
            int k = 0;
            int d = 0;
            int ascii;
            string indexHex = hexChunk[3] + hexChunk[2] + hexChunk[1] + hexChunk[0];
            if (indexHex.Length == 0)
            {
                MessageBox.Show("Error: Invalid file format. IndexHex");
                Application.Current.Shutdown();
            }
            string timeStampHex = hexChunk[7] + hexChunk[6] + hexChunk[5] + hexChunk[4];




            int num = Convert.ToInt32(indexHex, 16);
            int num1 = Convert.ToInt32(timeStampHex, 16);
            if (datIndexArray.Count == 0)
            {
                datIndexArray.Add(num);
            }
            else if (datIndexArray[datIndexArray.Count - 1] + 1 != num)
            {
                MessageBox.Show("Error: DatIndex must be sequential");
                Application.Current.Shutdown();
            }
            else
            {
                datIndexArray.Add(num);
            }

            if (timeArray.Count == 0)
            {
                timeArray.Add(num1);
            }
            else if (timeArray[timeArray.Count - 1] > num1)
            {
                MessageBox.Show("Error: Time value must be greater than the previous values.");
                Application.Current.Shutdown();
            }
            else
            {
                timeArray.Add(num1);
            }



            for (int i = 8; i < hexChunk.Length - 5; i += 4)
            {
                string x = hexChunk[i + 3] + hexChunk[i + 2] + hexChunk[i + 1] + hexChunk[i];
                int decimalValue = Convert.ToInt32(x, 16);

                string binaryValue = Convert.ToString(decimalValue, 2).PadLeft(32, '0');

                if (binaryValue[0] == '1')
                {
                    string invertedBits = "";
                    for (int j = 0; j < binaryValue.Length; j++)
                    {
                        invertedBits += binaryValue[j] == '1' ? '0' : '1'; // Finding One's Complement
                    }
                    int a = Convert.ToInt32(invertedBits, 2);
                    int ans = (a + 1) * (-1);
                    ascii = ans;
                }
                else
                {
                    ascii = decimalValue;
                }



                double result = 0;
                if (Analog[k].DataPrimarySecondary.Equals("S"))
                {

                    result = ((ascii * Analog[k].ChannelMultiplier) + Analog[k].ChannelOffset) * Analog[k].ChannelRatioPrimary;
                }

                else
                {
                    result = (ascii * Analog[k].ChannelMultiplier) + Analog[k].ChannelOffset;
                }

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    using (SqlTransaction transaction = con.BeginTransaction())
                    {


                        using (SqlCommand cmdAnalogDat = new SqlCommand(query4, con, transaction))
                        {
                            cmdAnalogDat.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                            cmdAnalogDat.Parameters.AddWithValue("@AnalogIndex", Analog[k].ChannelIndexNumber);
                            cmdAnalogDat.Parameters.AddWithValue("@DatIndex", num);
                            cmdAnalogDat.Parameters.AddWithValue("@Time", num1);
                            cmdAnalogDat.Parameters.AddWithValue("@Value", ascii);
                            cmdAnalogDat.Parameters.AddWithValue("@Result", result);

                            cmdAnalogDat.ExecuteNonQuery();
                        }

                        transaction.Commit();
                    }
                }
                k++;








            }

            for (int i = 8 + (4 * AnalogCount); i < hexChunk.Length; i = i + 2)
            {

                string statusBitsHex = hexChunk[i + 1] + hexChunk[i];
                int bits = Convert.ToInt32(statusBitsHex, 16);
                string statusBitsBin = Convert.ToString(bits, 2).PadLeft(16, '0');
                for (int j = statusBitsBin.Length - 1; j >= 0; j--)
                {

                    using (SqlConnection con = new SqlConnection(connectionString))
                    {
                        con.Open();
                        using (SqlTransaction transaction = con.BeginTransaction())
                        {
                            using (SqlCommand cmdDigitalDat = new SqlCommand(query5, con, transaction))
                            {
                                cmdDigitalDat.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                                cmdDigitalDat.Parameters.AddWithValue("@DigitalIndex", Digital[d].ChannelNumber);
                                cmdDigitalDat.Parameters.AddWithValue("@DatIndex", num);
                                cmdDigitalDat.Parameters.AddWithValue("@Time", num1);
                                cmdDigitalDat.Parameters.AddWithValue("@Value", statusBitsBin[j]);
                                cmdDigitalDat.ExecuteNonQuery();
                            }
                            transaction.Commit();
                            con.Close();
                        }
                    }
                    d++;

                }


            }
        }

        //Function for float32




        public static void Float32(string[] hexChunk, int AnalogCount, int DigitalCount)
        {

            string connectionString = "Data Source=SANAL-PROEDISON\\SQLEXPRESS;Initial Catalog=Demo;User ID=sa;Password=mypassword;Encrypt=False;";
            string query4 = "INSERT INTO AnalogDat(ComtradeIndex,AnalogIndex,DatIndex,Time,Value,Result) VALUES (@ComtradeIndex,@AnalogIndex,@DatIndex,@Time,@Value,@Result)";
            string query5 = "INSERT INTO DigitalDat(ComtradeIndex,DigitalIndex,DatIndex,Time,Value) VALUES (@ComtradeIndex,@DigitalIndex,@DatIndex,@Time,@Value)";
            int k = 0;
            int d = 0;
            string indexHex = hexChunk[3] + hexChunk[2] + hexChunk[1] + hexChunk[0];
            if (indexHex.Length == 0)
            {
                MessageBox.Show("Error: Invalid file format. IndexHex");
                Application.Current.Shutdown();
            }

            string timeStampHex = hexChunk[7] + hexChunk[6] + hexChunk[5] + hexChunk[4];

            int num = Convert.ToInt32(indexHex, 16);
            int num1 = Convert.ToInt32(timeStampHex, 16);
            if (datIndexArray.Count == 0)
            {
                datIndexArray.Add(num);
            }
            else if (datIndexArray[datIndexArray.Count - 1] + 1 != num)
            {
                MessageBox.Show("Error: DatIndex must be sequential");
                Application.Current.Shutdown();
            }
            else
            {
                datIndexArray.Add(num);
            }

            if (timeArray.Count == 0)
            {
                timeArray.Add(num1);
            }
            else if (timeArray[timeArray.Count - 1] > num1)
            {
                MessageBox.Show("Error: Time value must be greater than the previous values.");
                Application.Current.Shutdown();
            }
            else
            {
                timeArray.Add(num1);
            }




            for (int i = 8; i < hexChunk.Length - 5; i += 4)
            {
                string x = hexChunk[i + 3] + hexChunk[i + 2] + hexChunk[i + 1] + hexChunk[i];
                int decimalValue = Convert.ToInt32(x, 16);

                string binaryValue = Convert.ToString(decimalValue, 2).PadLeft(32, '0');

                double mantisa = ConvertBinaryFractionToDecimal(binaryValue.Substring(9)) + 1.0;
                int exponent = Convert.ToInt32(binaryValue.Substring(1, 8), 2) - 127;

                double ascii = mantisa * Math.Pow(2, exponent);


                if (binaryValue[0] == '1')
                {
                    ascii = -ascii;
                }
                double result = 0.0;
                if (Analog[k].DataPrimarySecondary.Equals("S"))
                {

                    result = ((ascii * Analog[k].ChannelMultiplier) + Analog[k].ChannelOffset) * Analog[k].ChannelRatioPrimary;
                }

                else
                {
                    result = (ascii * Analog[k].ChannelMultiplier) + Analog[k].ChannelOffset;
                }




                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    using (SqlTransaction transaction = con.BeginTransaction())
                    {


                        using (SqlCommand cmdAnalogDat = new SqlCommand(query4, con, transaction))
                        {
                            cmdAnalogDat.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                            cmdAnalogDat.Parameters.AddWithValue("@AnalogIndex", Analog[k].ChannelIndexNumber);
                            cmdAnalogDat.Parameters.AddWithValue("@DatIndex", num);
                            cmdAnalogDat.Parameters.AddWithValue("@Time", num1);
                            cmdAnalogDat.Parameters.AddWithValue("@Value", ascii);
                            cmdAnalogDat.Parameters.AddWithValue("@Result", result);

                            cmdAnalogDat.ExecuteNonQuery();
                        }

                        transaction.Commit();
                    }
                }
                k++;



            }
            for (int i = 8 + (4 * AnalogCount); i < hexChunk.Length; i = i + 2)
            {

                string statusBitsHex = hexChunk[i + 1] + hexChunk[i];
                int bits = Convert.ToInt32(statusBitsHex, 16);
                string statusBitsBin = Convert.ToString(bits, 2).PadLeft(16, '0');
                for (int j = statusBitsBin.Length - 1; j >= 0; j--)
                {

                    using (SqlConnection con = new SqlConnection(connectionString))
                    {
                        con.Open();
                        using (SqlTransaction transaction = con.BeginTransaction())
                        {
                            using (SqlCommand cmdDigitalDat = new SqlCommand(query5, con, transaction))
                            {
                                cmdDigitalDat.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                                cmdDigitalDat.Parameters.AddWithValue("@DigitalIndex", Digital[d].ChannelNumber);
                                cmdDigitalDat.Parameters.AddWithValue("@DatIndex", num);
                                cmdDigitalDat.Parameters.AddWithValue("@Time", num1);
                                cmdDigitalDat.Parameters.AddWithValue("@Value", statusBitsBin[j]);
                                cmdDigitalDat.ExecuteNonQuery();
                            }
                            transaction.Commit();
                            con.Close();
                        }
                    }
                    d++;

                }


            }



        }


        static double ConvertBinaryFractionToDecimal(string binaryFraction)
        {
            double decimalValue = 0;

            for (int i = 0; i < binaryFraction.Length; i++)
            {
                int bit = binaryFraction[i] - '0'; // Convert char ('1' or '0') to int
                double fractionalValue = bit * Math.Pow(2, -(i + 1));
                decimalValue += fractionalValue;
            }

            return decimalValue;
        }








        // Handle "Choose File" button click
        private void btnChooseFile_Click(object sender, RoutedEventArgs e)
        {

            // Open file dialog for selecting a file
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select a File",
                Filter = "CFG Files (*.cfg)|*.cfg|DAT Files (*.dat)|*.dat" // Filter for cfg files
            };

            // Show the dialog and check if the user selected a file
            if (openFileDialog.ShowDialog() == true)
            {
                filePath = openFileDialog.FileName;
                fileName = Path.GetFileName(filePath);
                fileExtension = Path.GetExtension(filePath).ToLower();
            }
        }

        // Handle "Process File" button click
        private void btnProcessFile_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Processing file...");
            // Check if a file has been selected
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("No file selected.");
                return;
            }

            var fileLines = File.ReadAllLines(filePath);

            int analogIndex = 0, digitalIndex = 0;
            if (string.Equals(fileExtension, ".cfg", StringComparison.OrdinalIgnoreCase))
            {
                Words.Clear();


                // Read and process lines
                string line = fileLines[0];
                var tokens = line.Split(',');
                int year = int.Parse(tokens[2]);


                if (year == 2013)
                {
                    ExtractRevisedYear(line);
                    line = fileLines[1];
                    SignalCounting(line);

                    if (fileLines.Length != (11 + Comtrade1.AnalogSignalCount + Comtrade1.DigitalSignalCount))
                    {
                        string n = Convert.ToString(fileLines.Length);
                        MessageBox.Show("Error: Invalid file format. Total lines");
                        Application.Current.Shutdown();
                    }

                    for (int i = 2; i < Comtrade1.AnalogSignalCount + 2; i++)
                    {
                        ParseAndStoreAnalogData(fileLines[i]);
                        analogIndex++;
                    }
                    for (int i = Comtrade1.AnalogSignalCount + 2; i < Comtrade1.DigitalSignalCount + Comtrade1.AnalogSignalCount + 2; i++)
                    {
                        ParseAndStoreDigitalData(fileLines[i]);
                    }
                    for (int i = Comtrade1.AnalogSignalCount + Comtrade1.DigitalSignalCount + 2; i < fileLines.Length; i++)
                    {
                        ComtradeParse(fileLines[i]);
                    }
                    ProcessWords(Words);
                }

                else
                {

                    ExtractRevisedYear1999(fileLines[0]);

                    SignalCounting(fileLines[1]);

                    for (int i = 2; i < Comtrade1.AnalogSignalCount + 2; i++)
                    {
                        ParseAndStoreAnalogData1999(fileLines[i]);
                        analogIndex++;
                    }

                    for (int i = Comtrade1.AnalogSignalCount + 2; i < Comtrade1.DigitalSignalCount + Comtrade1.AnalogSignalCount + 2; i++)
                    {
                        ParseAndStoreDigitalData1999(fileLines[i]);
                    }
                    for (int i = Comtrade1.AnalogSignalCount + Comtrade1.DigitalSignalCount + 2; i < fileLines.Length; i++)
                    {
                        ComtradeParse1999(fileLines[i]);
                    }
                    ProcessWords1999(Words);


                }

            }

            if (string.Equals(fileExtension, ".dat", StringComparison.OrdinalIgnoreCase))
            {
                timeArray.Clear();
                datIndexArray.Clear();

                MessageBox.Show("Dat file started");
                if (string.Equals(Comtrade.DataType, "ASCII", StringComparison.OrdinalIgnoreCase))
                {
                    timeArray.Clear();
                    datIndexArray.Clear();

                    var AsciiLines = File.ReadAllLines(filePath);
                    if (AsciiLines.Length != Comtrade.LastSampleRate)
                    {
                        MessageBox.Show("Error: Invalid file format. ASCII");
                        Application.Current.Shutdown();
                    }
                    using (StreamReader reader = new StreamReader(filePath))
                    {
                        string lineDat;
                        while ((lineDat = reader.ReadLine()) != null)
                        {
                            // Process each line
                            AsciiDat(lineDat, Comtrade1.AnalogSignalCount, Comtrade1.DigitalSignalCount);
                        }
                    }
                    MessageBox.Show("Ascii Completed");
                }
                else if (string.Equals(Comtrade.DataType, "BINARY", StringComparison.OrdinalIgnoreCase))
                {
                    timeArray.Clear();
                    datIndexArray.Clear();

                    if (!File.Exists(filePath))
                    {
                        MessageBox.Show("Error: File does not exist.");
                        return;
                    }

                    byte[] binaryData = File.ReadAllBytes(filePath);

                    // Convert binary data to hexadecimal and store in an array
                    string[] hexArray = new string[binaryData.Length];
                    for (int i = 0; i < binaryData.Length; i++)
                    {
                        hexArray[i] = $"{binaryData[i]:X2}"; // Convert each byte to a 2-digit hex string
                    }

                    // Define chunk size (Based on number of analog signals)
                    int chunkSize = 8 + (2 * Comtrade1.AnalogSignalCount) + (2 * (Comtrade1.DigitalSignalCount / 16));

                    for (int i = 0; i < hexArray.Length; i += chunkSize)
                    {
                        int currentChunkSize = Math.Min(chunkSize, hexArray.Length - i);
                        string[] currentChunk = new string[currentChunkSize];
                        Array.Copy(hexArray, i, currentChunk, 0, currentChunkSize);

                        // Call the processing function with the current chunk
                        BinaryDat(currentChunk, Comtrade1.AnalogSignalCount, Comtrade1.DigitalSignalCount);
                    }
                    MessageBox.Show("Binary Completed");
                }
                else if (string.Equals(Comtrade.DataType, "BINARY32", StringComparison.OrdinalIgnoreCase))
                {
                    timeArray.Clear();
                    datIndexArray.Clear();
                    if (!File.Exists(filePath))
                    {
                        MessageBox.Show("Error: File does not exist.");
                        return;
                    }

                    byte[] binaryData = File.ReadAllBytes(filePath);

                    // Convert binary data to hexadecimal and store in an array
                    string[] hexArray = new string[binaryData.Length];
                    for (int i = 0; i < binaryData.Length; i++)
                    {
                        hexArray[i] = $"{binaryData[i]:X2}"; // Convert each byte to a 2-digit hex string
                    }

                    // Define chunk size (Based on number of analog signals)
                    int chunkSize = 8 + (4 * Comtrade1.AnalogSignalCount) + (2 * (Comtrade1.DigitalSignalCount / 16));

                    for (int i = 0; i < hexArray.Length; i += chunkSize)
                    {
                        int currentChunkSize = Math.Min(chunkSize, hexArray.Length - i);
                        string[] currentChunk = new string[currentChunkSize];
                        Array.Copy(hexArray, i, currentChunk, 0, currentChunkSize);

                        // Call the processing function with the current chunk
                        Binary32Dat(currentChunk, Comtrade1.AnalogSignalCount, Comtrade1.DigitalSignalCount);
                    }
                    MessageBox.Show("Binary32 Completed");
                }
                else if (string.Equals(Comtrade.DataType, "FLOAT32", StringComparison.OrdinalIgnoreCase))
                {
                    timeArray.Clear();
                    datIndexArray.Clear();
                    if (!File.Exists(filePath))
                    {
                        MessageBox.Show("Error: File does not exist.");
                        return;
                    }
                    byte[] binaryData = File.ReadAllBytes(filePath);

                    // Convert binary data to hexadecimal and store in an array
                    string[] hexArray = new string[binaryData.Length];
                    for (int i = 0; i < binaryData.Length; i++)
                    {
                        hexArray[i] = $"{binaryData[i]:X2}"; // Convert each byte to a 2-digit hex string
                    }

                    // Define chunk size (Based on number of analog signals)
                    int chunkSize = 8 + (4 * Comtrade1.AnalogSignalCount) + (2 * (Comtrade1.DigitalSignalCount / 16));

                    for (int i = 0; i < hexArray.Length; i += chunkSize)
                    {
                        int currentChunkSize = Math.Min(chunkSize, hexArray.Length - i);
                        string[] currentChunk = new string[currentChunkSize];
                        Array.Copy(hexArray, i, currentChunk, 0, currentChunkSize);

                        // Call the processing function with the current chunk
                        Float32(currentChunk, Comtrade1.AnalogSignalCount, Comtrade1.DigitalSignalCount);
                    }
                    MessageBox.Show("Float32 Completed");
                }
                else
                {
                    MessageBox.Show($"Unknown data type: {Comtrade.DataType}");
                }





            }

            string connectionString = "Data Source=SANAL-PROEDISON\\SQLEXPRESS;Initial Catalog=Demo;User ID=sa;Password=mypassword;Encrypt=False;";

            // Queries
            string query = "INSERT INTO COMTRADE (ComtradeIndex,FilePath,FileName,Error,isProedison,HasHDR) VALUES (@ComtradeIndex,@FilePath,@FileName,@Error,@isProedison,@HasHDR)";
            string query1 = "INSERT INTO Analog(ComtradeIndex,AnalogIndex,ChannelID,Phase,CCBM,Units,Multiplier,Offset,Skew,minVal,maxVal,PrimaryVal,SecondaryVal,ChannelType) VALUES (@ComtradeIndex,@AnalogIndex,@ChannelID,@Phase,@CCBM,@Units,@Multiplier,@Offset,@Skew,@minVal,@maxVal,@PrimaryVal,@SecondaryVal,@ChannelType)";
            string query2 = "INSERT INTO Digital(ComtradeIndex,DigitalIndex,ChannelID,Phase,CCBM,InitialState) VALUES (@ComtradeIndex,@DigitalIndex,@ChannelID,@Phase,@CCBM,@InitialState)";
            string query3 = "INSERT INTO CFG(ComtradeIndex,Station,DeviceID,CfgVersion,Frequency,SampleRate,SampleCountHz,LastSampleCount,FirstSampleTime,TriggerTime,DataType,TimeMultiplier,LocalTime,UTCTime,TimeQualityIndicatorCode,LeapSecond) VALUES (@ComtradeIndex,@Station,@DeviceID,@CfgVersion,@Frequency,@SampleRate,@SampleCountHz,@LastSampleCount,@FirstSampleTime,@TriggerTime,@DataType,@TimeMultiplier,@LocalTime,@UTCTime,@TimeQualityIndicatorCode,@LeapSecond)";
            string query4 = "INSERT INTO AnalogDat(ComtradeIndex,ChannelIndex,DatIndex,Time,Value) VALUES (@ComtradeIndex,@ChannelIndex,@DatIndex,@Time,@Value)";


            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                if (string.Equals(fileExtension, ".cfg", StringComparison.OrdinalIgnoreCase))
                {
                    using (SqlTransaction transaction = con.BeginTransaction())
                    {
                        ComtradeIndex = 0;

                        string countQuery = "SELECT COUNT(*) FROM COMTRADE";
                        using (SqlCommand countCmd = new SqlCommand(countQuery, con, transaction))
                        {
                            ComtradeIndex = (int)countCmd.ExecuteScalar() + 1;
                        }

                        // Insert into Comtrade
                        using (SqlCommand cmd = new SqlCommand(query, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                            cmd.Parameters.AddWithValue("@FilePath", filePath);
                            cmd.Parameters.AddWithValue("@FileName", fileName);
                            cmd.Parameters.AddWithValue("@Error", 0); // Assuming no error, adjust based on your needs
                            cmd.Parameters.AddWithValue("@isProedison", true); // Adjust as per your requirement
                            cmd.Parameters.AddWithValue("@HasHDR", false); // Adjust as per your requirement

                            cmd.ExecuteNonQuery();
                        }

                        // Insert into AnalogData
                        using (SqlCommand cmdAnalog = new SqlCommand(query1, con, transaction))
                        {
                            foreach (var analog in Analog)
                            {
                                cmdAnalog.Parameters.Clear();
                                cmdAnalog.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                                cmdAnalog.Parameters.AddWithValue("@AnalogIndex", analog.ChannelIndexNumber);
                                cmdAnalog.Parameters.AddWithValue("@ChannelID", analog.ChannelId);
                                cmdAnalog.Parameters.AddWithValue("@Phase", analog.PhaseId);
                                cmdAnalog.Parameters.AddWithValue("@CCBM", analog.Ccbm);
                                cmdAnalog.Parameters.AddWithValue("@Units", analog.ChannelUnits);
                                cmdAnalog.Parameters.AddWithValue("@Multiplier", analog.ChannelMultiplier);
                                cmdAnalog.Parameters.AddWithValue("@Offset", analog.ChannelOffset);
                                cmdAnalog.Parameters.AddWithValue("@Skew", analog.ChannelSkew);
                                cmdAnalog.Parameters.AddWithValue("@minVal", analog.MinimumLimit);
                                cmdAnalog.Parameters.AddWithValue("@maxVal", analog.MaximumLimit);
                                cmdAnalog.Parameters.AddWithValue("@PrimaryVal", analog.ChannelRatioPrimary);
                                cmdAnalog.Parameters.AddWithValue("@SecondaryVal", analog.ChannelRatioSecondary);
                                cmdAnalog.Parameters.AddWithValue("@ChannelType", analog.DataPrimarySecondary);

                                cmdAnalog.ExecuteNonQuery();
                            }
                        }

                        // Insert into DigitalData
                        using (SqlCommand cmdDigital = new SqlCommand(query2, con, transaction))
                        {
                            foreach (var digital in Digital)
                            {
                                cmdDigital.Parameters.Clear();
                                cmdDigital.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                                cmdDigital.Parameters.AddWithValue("@DigitalIndex", digital.ChannelNumber);
                                cmdDigital.Parameters.AddWithValue("@ChannelID", digital.ChannelId);
                                cmdDigital.Parameters.AddWithValue("@Phase", digital.PhaseId);
                                cmdDigital.Parameters.AddWithValue("@CCBM", digital.Ccbm);
                                cmdDigital.Parameters.AddWithValue("@InitialState", digital.NormalState);

                                cmdDigital.ExecuteNonQuery();
                            }
                        }

                        // Insert into CFG
                        // Insert into CFG
                        using (SqlCommand cmdcfg = new SqlCommand(query3, con, transaction))
                        {
                            cmdcfg.Parameters.AddWithValue("@ComtradeIndex", ComtradeIndex);
                            cmdcfg.Parameters.AddWithValue("@Station", Comtrade.Station);
                            cmdcfg.Parameters.AddWithValue("@DeviceID", Comtrade.DeviceId);
                            cmdcfg.Parameters.AddWithValue("@CfgVersion", Comtrade.CfgVersion);
                            cmdcfg.Parameters.AddWithValue("@Frequency", Comtrade.LineFrequency);
                            cmdcfg.Parameters.AddWithValue("@SampleRate", Comtrade.SampleRate);
                            cmdcfg.Parameters.AddWithValue("@SampleCountHz", Comtrade.SampleRateCount);
                            cmdcfg.Parameters.AddWithValue("@LastSampleCount", Comtrade.LastSampleRate);
                            cmdcfg.Parameters.AddWithValue("@FirstSampleTime", Comtrade.FirstTimeStamp);
                            cmdcfg.Parameters.AddWithValue("@TriggerTime", Comtrade.TriggerTimeStamp);
                            cmdcfg.Parameters.AddWithValue("@DataType", Comtrade.DataType);
                            cmdcfg.Parameters.AddWithValue("@TimeMultiplier", Comtrade.TimeMultiplier);
                            cmdcfg.Parameters.AddWithValue("@LocalTime", (object)Comtrade.TimeCode ?? DBNull.Value);
                            cmdcfg.Parameters.AddWithValue("@UTCTime", (object)Comtrade.LocalCode ?? DBNull.Value);
                            cmdcfg.Parameters.AddWithValue("@TimeQualityIndicatorCode", (object)Comtrade.TimeQualityIndicatorCode ?? DBNull.Value);
                            cmdcfg.Parameters.AddWithValue("@LeapSecond", (object)Comtrade.LeapSecondIndicator ?? DBNull.Value);

                            cmdcfg.ExecuteNonQuery();
                        }
                        // Commit the transaction
                        transaction.Commit();
                        MessageBox.Show("Completed");
                    }
                }







            }
        }
    }
}





