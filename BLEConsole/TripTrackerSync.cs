using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Security.Cryptography;

namespace BLEConsole
{
    internal static class TripTrackerSync
    {
        enum PendingWorkType { None, RequestData, ProcessResults, Done };
        static PendingWorkType PendingWork { get; set; } = PendingWorkType.None;

        static bool _aggregrateDataInValueChanged = false;
        static int _aggregrateDataLengthRemaining = 0;
        static byte[] _aggregateDataArray = null;
        static int StartId { get; set; } = 0;
        static int EndId { get; set; } = 0;

        public static bool Debug { get; set; } = false;
        public static List<Event2Info> EventInfoList { get; set; } = new List<Event2Info>();

        public static async Task<int> Initialize(string deviceName)
        {
            int result = 0;
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
                    // request starting and ending Ids of the events to sync
                    // result will be processed in the Characteristic_ValueChanged Handler
                    result += await Program.WriteCharacteristic("#0 SC?");
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
                        await RequestEventData();
                        break;

                    case PendingWorkType.ProcessResults:
                        PendingWork = PendingWorkType.Done;
                        await ProcessEventData();
                        break;

                    default:
                        Thread.Sleep(200);
                        break;
                }
            }
        }

        public static async Task RequestEventData()
        {
            int result = 0;

            try
            {
                if (Debug) Console.WriteLine($"\nSyncRange {StartId} - {EndId}");

                for (int i = StartId; i < EndId; i++)
                {
                    string cmd = $"#0 SD?{i};";
                    if (Debug) Console.WriteLine($"\nEvent request {i} {cmd}");
                    result += await Program.WriteCharacteristic(cmd);
                    Thread.Sleep(200);
                    if (result > 0) break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"RequestEventData: Exception {ex}");
                result++;
            }
            if (result > 0)
            {
                Console.WriteLine($"RequestEventData failed");
            }
        }

        public static async Task ProcessEventData()
        {
            int result = 0;

            // Store all the Trip Tracker data for the day into an existing
            // Excel spreadsheet
            // TODO: make the spreadsheet be setable rather than a constant.
            var filename = "C:\\Users\\drmcl\\GitHub\\Temp\\FR3-RV-Log.xlsm";
            Console.Write($"Writing data to Excel: {filename} ... ");

            ExcelWriter eWriter = new ExcelWriter(filename);
            bool Saved = eWriter.AppendToTripDetailTable(EventInfoList);
            eWriter.Dispose();

            if ( Saved ) {
                Console.WriteLine("Complete");

                // Save was successful so update the SyncStartId
                try
                {
                    result += await Program.WriteCharacteristic($"#0 SS={EndId};");
                    Thread.Sleep(200);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"RequestEventData: Exception {ex}");
                    result++;
                }
                if (result > 0)
                {
                    Console.WriteLine($"RequestEventData failed");
                }
            } 
            else
            {
                Console.WriteLine("Write Failed");
            }
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

            if (tempValue.StartsWith("SyncRange="))
            {
                // Format is 'SyncRange=<StartId>,<EndId>\n'
                var parts = tempValue.Split('=');
                var parts2 = parts[1].Split('\n');
                var counts = parts2[0].Split(',');
                StartId = Convert.ToInt32(counts[0], 10);
                EndId = Convert.ToInt32(counts[1], 10);

                if (EndId > StartId)
                {
                    PendingWork = PendingWorkType.RequestData;
                }
                else
                {
                    Console.Write($"SyncRange is Empty - Aborting\n");
                    PendingWork = PendingWorkType.Done;
                }
                
                if (Debug) Console.Write($"\nValue changed for {sender.Uuid} processing {parts[0]}={parts2[0]}\n");

                processed = true;
            }
            else if (tempValue.StartsWith("SyncData="))
            {
                startAggregation = true;
                processed = true;
            }

            if (startAggregation)
            {
                // The format of the command is SyncData=<hex Length of binary data>\n<start of binary data>
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

                    // The 8th byte of the data is the type of the Event.
                    Event2Info.EventType type = (Event2Info.EventType)_aggregateDataArray[8];

                    switch (type)
                    {
                        case Event2Info.EventType.StartDay:
                            var sdp = new StartDayInfo(_aggregateDataArray);
                            EventInfoList.Add(sdp);
                            sdp.Print();
                            break;

                        case Event2Info.EventType.StartLeg:
                            var slp = new StartLegInfo(_aggregateDataArray);
                            EventInfoList.Add(slp);
                            slp.Print();
                            break;

                        case Event2Info.EventType.EndLeg:
                            var elp = new EndLegInfo(_aggregateDataArray);
                            EventInfoList.Add(elp);
                            elp.Print();
                            break;

                        case Event2Info.EventType.EndDay:
                            var edp = new EndDayInfo(_aggregateDataArray);
                            EventInfoList.Add(edp);
                            edp.Print();
                            break;

                        case Event2Info.EventType.Gas:
                            var gas = new GasInfo(_aggregateDataArray);
                            EventInfoList.Add(gas);
                            gas.Print();
                            break;

                        case Event2Info.EventType.Propane:
                            var propane = new PropaneInfo(_aggregateDataArray);
                            EventInfoList.Add(propane);
                            propane.Print();
                            break;

                        case Event2Info.EventType.OilChange:
                            var oilChange = new OilChangeInfo(_aggregateDataArray);
                            EventInfoList.Add(oilChange);
                            oilChange.Print();
                            break;
                        
                        default:
                            Console.WriteLine($"ERROR: Aggregation of EventInfo got invalid type '{_aggregateDataArray[0]}");
                            break;
                    }
                    // we have all the requested data 
                    // day summary is first, and then NumLegs 
                    // and we have all the Event info
                    if (EventInfoList.Count >= (EndId - StartId))
                    {
                        PendingWork = PendingWorkType.ProcessResults;
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
