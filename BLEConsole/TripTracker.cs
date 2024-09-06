using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Printers;
using Windows.Security.Cryptography;
using Windows.UI.Xaml.Controls.Maps;
using Excel = Microsoft.Office.Interop.Excel;

namespace BLEConsole
{
    public static class TripTracker
    {
        enum AggregateDataType { None, VehicleInfo, TripInfo, EventInfo };
        enum PendingWorkType { None, RequestData, ProcessResults, Done };

        static bool _aggregrateDataInValueChanged = false;
        static int _aggregrateDataLengthRemaining = 0;
        static byte[] _aggregateDataArray = null;
        static AggregateDataType _aggregateDataType = AggregateDataType.None;
        static PendingWorkType PendingWork { get; set; } = PendingWorkType.None;
        static int NumLegs { get; set; } = 0;
        static int NumEvents { get; set; } = 0;

        public static bool Debug { get; set; } = false;

        public static List<TripInfo> TripInfoList { get; set; } = new List<TripInfo>();
        public static VehicleInfo VehicleInfo { get; set; }
        public static List<EventInfo> EventInfoList { get; set; } = new List<EventInfo>();


        public static async Task<int> Initialize(string deviceName)
        {
            int result = 0;
            TripInfoList = new List<TripInfo>();
            EventInfoList = new List<EventInfo>();

            try
            {
                result += await Program.OpenDevice(deviceName);

                if (result == 0)
                {
                    result += await Program.SetService("SimpleKeyService");
                }

                if (result == 0)
                {
                    result += await Program.SubscribeToCharacteristic("#0"); // SimpleKeyState
                }

                // start the listener for processing the results
                PendingWork = PendingWorkType.None;

                if (result == 0)
                {
                    // request the number of Legs and Events in the current Day
                    // result will be processed in the Characteristic_ValueChanged Handler
                    result += await Program.WriteCharacteristic("#0 LE?");
                    Thread.Sleep(200);
                }
                if (result == 0)
                {
                    await ProcessPendingWork();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                result++;
            }
            return result;
        }

        public static async Task ProcessPendingWork()
        {
            while (PendingWork != PendingWorkType.Done)
            {

                if (Debug) Console.WriteLine($"ProcessPendingWork {PendingWork}");

                switch (PendingWork)
                {
                    case PendingWorkType.RequestData:
                        PendingWork = PendingWorkType.None;
                        await RequestTripData();
                        break;

                    case PendingWorkType.ProcessResults:
                        PendingWork = PendingWorkType.Done;
                        ProcessTripData();
                        break;

                    default:
                        Thread.Sleep(200);
                        break;
                }
            }
        }

        public static async Task RequestTripData()
        {
            int result = 0;

            try
            {
                if (Debug) Console.WriteLine($"\nTripData {NumLegs}\nrequest VH?");

                result += await Program.WriteCharacteristic("#0 VH?");
                Thread.Sleep(200);

                if (result == 0)
                {
                    if (Debug) Console.WriteLine("\nrequest DD?");
                    result += await Program.WriteCharacteristic("#0 DD?");
                    Thread.Sleep(200);
                }

                if (result == 0)
                {
                    for (int i = 0; i < NumLegs; i++)
                    {
                        string cmd = $"#0 LD?{i};";
                        if (Debug) Console.WriteLine($"\nLeg request {i} {cmd}");
                        result += await Program.WriteCharacteristic(cmd);
                        Thread.Sleep(200);
                        if (result > 0) break;
                    }
                }

                if (result == 0)
                {
                    for (int i = 0; i < NumEvents; i++)
                    {
                        string cmd = $"#0 ED?{i};";
                        if (Debug) Console.WriteLine($"\nEvent request {i} {cmd}");
                        result += await Program.WriteCharacteristic(cmd);
                        Thread.Sleep(200);
                        if (result > 0) break;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"RequestTripData: Exception {ex}");
                result++;
            }
            if (result > 0)
            {
                Console.WriteLine($"RequestTripData failed");
            }
        }

        enum ExcelInsertType { DayStart, LegStart, LegEnd, DayEnd };

        public static void ProcessTripData()
        {
            // Store all the Trip Tracker data for the day into an existing
            // Excel spreadsheet
            // TODO: make the spreadsheet be setable rather than a constant.
            var filename = "C:\\Users\\drmcl\\GitHub\\Temp\\FR3-RV-Log.xlsm";
            Console.Write($"Writing data to Excel: {filename} ... ");

            ExcelWriter eWriter = new ExcelWriter(filename);
            eWriter.AppendToTripDetailTable(VehicleInfo, TripInfoList, EventInfoList);
            eWriter.Dispose();

            Console.WriteLine("Complete");
        }

        /// <summary>
        /// Event handler for ValueChanged callback
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public static bool Characteristic_ValueChanged(GattCharacteristic sender, GattValueChangedEventArgs args)
        {
            var tempValue = Utilities.FormatValue(args.CharacteristicValue, DataFormat.UTF8);

            bool startAggregation = false;
            bool processed = false;

            if (tempValue.StartsWith("LegEventCounts="))
            {
                // Format is 'LegEventCounts=<NumLegs>,<NumEvents>\n<start of binary data>'
                var parts = tempValue.Split('=');
                var parts2 = parts[1].Split('\n');
                var counts = parts2[0].Split(',');
                NumLegs = Convert.ToInt32(counts[0], 10);
                NumEvents = Convert.ToInt32(counts[1], 10);

                PendingWork = PendingWorkType.RequestData;
                if (Debug) Console.Write($"\nValue changed for {sender.Uuid} processing {parts[0]}={parts2[0]}\n");

                processed = true;
            }
            else if (tempValue.StartsWith("VH="))
            {
                _aggregateDataType = AggregateDataType.VehicleInfo;
                startAggregation = true;
                processed = true;
            }
            else if (tempValue.StartsWith("TP="))
            {
                _aggregateDataType = AggregateDataType.TripInfo;
                startAggregation = true;
                processed = true;
            }
            else if (tempValue.StartsWith("ED="))
            {
                _aggregateDataType = AggregateDataType.EventInfo;
                startAggregation = true;
                processed = true;
            }

            if (startAggregation)
            {
                // The format of the command is XX=<hex Length of binary data>\n<start of binary data>
                var parts = tempValue.Split('=');
                var parts2 = parts[1].Split('\n');
                _aggregrateDataLengthRemaining = Convert.ToInt32(parts2[0], 16);
                if (Debug) Console.Write($"\nValue changed for {sender.Uuid} processing {parts[0]}={parts2[0]}\n");

                parts = tempValue.Split('\n');

                byte[] data;
                CryptographicBuffer.CopyToByteArray(args.CharacteristicValue, out data);
                _aggregateDataArray = data.Skip(parts[0].Length + 1).ToArray();
                _aggregrateDataLengthRemaining -= _aggregateDataArray.Length;

                if (_aggregrateDataLengthRemaining > 0)
                    _aggregrateDataInValueChanged = true;

                if (Debug)
                {
                    if (Console.IsInputRedirected) Console.Write($"{parts[0]}");
                    else Console.Write($"Value changed for {sender.Uuid} ({parts[0].Length} bytes):\n{parts[0]}\n*** Aggregating data (captured {parts2[1].Length} bytes {_aggregrateDataLengthRemaining} bytes remaining) ***\nBLE: ");
                }

                processed = true;
            }
            else if (_aggregrateDataInValueChanged)
            {
                byte[] data;
                CryptographicBuffer.CopyToByteArray(args.CharacteristicValue, out data);

                _aggregateDataArray = _aggregateDataArray.Concat(data).ToArray();
                _aggregrateDataLengthRemaining -= (int)args.CharacteristicValue.Length;

                if (_aggregrateDataLengthRemaining <= 0)
                {
                    _aggregrateDataInValueChanged = false;

                    if (_aggregateDataType == AggregateDataType.VehicleInfo)
                    {
                        VehicleInfo = new VehicleInfo(_aggregateDataArray);
                        Console.WriteLine("Vehicle Info");
                        VehicleInfo.Print();
                    }
                    else if (_aggregateDataType == AggregateDataType.TripInfo)
                    {
                        var info = new TripInfo(_aggregateDataArray);
                        TripInfoList.Add(info);
                        if (info.Id == 0)
                        {
                            Console.WriteLine($"Day Info");
                        } 
                        else
                        {
                            Console.WriteLine($"Trip Leg {info.Id} Info");
                        }
                        info.Print();

                        // we have all the requested data 
                        // day summary is first, and then NumLegs 
                        // and we have all the Event info
                        if (TripInfoList.Count >= NumLegs + 1 && EventInfoList.Count >= NumEvents)
                        {
                            PendingWork = PendingWorkType.ProcessResults;
                        }
                    }
                    else if (_aggregateDataType == AggregateDataType.EventInfo)
                    {
                        // The first byte of the data is the type of the Event.
                        EventInfo.EventType type = (EventInfo.EventType)_aggregateDataArray[0];
                        switch (type)
                        {
                            case EventInfo.EventType.FuelPurchase:
                                var fp = new PurchaseFuelInfo(_aggregateDataArray);
                                EventInfoList.Add(fp);
                                fp.Print();
                                break;

                            case EventInfo.EventType.PropanePurchase:
                                var pp = new PurchasePropaneInfo(_aggregateDataArray);
                                EventInfoList.Add(pp);
                                pp.Print();
                                break;

                            case EventInfo.EventType.OilChange:
                                var oc = new ChangeOilInfo(_aggregateDataArray);
                                EventInfoList.Add(oc);
                                oc.Print();
                                break;

                            default:
                                Console.WriteLine($"ERROR: Aggregation of EventInfo got invalid type '{_aggregateDataArray[0]}");
                                break;
                        }
                        // we have all the requested data 
                        // day summary is first, and then NumLegs 
                        // and we have all the Event info
                        if (TripInfoList.Count >= NumLegs + 1 && EventInfoList.Count >= NumEvents)
                        {
                            PendingWork = PendingWorkType.ProcessResults;
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Error: Unknown aggregation type {_aggregateDataType}");
                    }
                }
                else
                {
                    if (Debug && !Console.IsInputRedirected) Console.Write($"Value changed for {sender.Uuid} ({args.CharacteristicValue.Length} bytes)\n *** Aggregating data (captured {_aggregateDataArray.Length} bytes {_aggregrateDataLengthRemaining} bytes remaining) ***\nBLE: ");
                }
                processed = true;
            }
            return processed;
        }
    }
}