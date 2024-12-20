using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Security.Cryptography;

namespace BLEConsole
{
    internal class TripTrackerSync : ExtensionBase
    {
        enum PendingWorkType { None, RequestData, ProcessResults, Done };
        private PendingWorkType PendingWork { get; set; } = PendingWorkType.None;

        bool _aggregrateDataInValueChanged = false;
        int _aggregrateDataLengthRemaining = 0;
        byte[] _aggregateDataArray = null;
        uint _lastIdprocessed = 0;
        string _excelFilename = "C:\\Users\\drmcl\\GitHub\\Temp\\FR3-RV-Log.xlsm";

        public int StartId { get; private set; } = 0;
        public int EndId { get; private set; } = 0;

        public bool Debug { get; set; } = false;
        public List<Event2Info> EventInfoList { get; private set; } = new List<Event2Info>();

        private (string filename, string parameters) ExtractFilename(string parameters)
        {
            string pattern = @"F:=""([^""]+)""|F:=([^ ]+)";
            var match = Regex.Match(parameters, pattern);

            if (match.Success)
            {
                string filename = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
                string modifiedString = parameters.Remove(match.Index, match.Length).Trim();
                return (filename, modifiedString);
            }

            return (null, parameters);
        }

        private bool IsFileWriteable(string path)
        {
            try
            {
                using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Write))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return false;
            }
        }

        public async Task<int> Initialize(string input)
        {
            var (filename, parameters) = ExtractFilename(input);
            if (filename != null)
            {
                _excelFilename = filename;
            }
            Console.WriteLine($"Will write the data to the excel file {_excelFilename} ");

            if (!IsFileWriteable(_excelFilename))
            {
                Console.WriteLine($"Aborting: {_excelFilename} not accessible");
                return 1;
            }

            int result = 0;
            try
            {
                result += await Program.OpenDevice(parameters);

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

        /// <summary>
        /// Async method to process work generated in Characteristic_ValueChanged
        /// </summary>
        private async Task ProcessPendingWork()
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

        /// <summary>
        /// Async method to request the Event data in the Sync range
        /// </summary>
        private async Task RequestEventData()
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
                    if (result > 0) break;
                    while (_lastIdprocessed < i)
                        Thread.Sleep(100);
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

        /// <summary>
        /// Async method to process the event data and write it to an excel file
        /// </summary>
        private async Task ProcessEventData()
        {
            int result = 0;

            // Store all the Trip Tracker data for the day into an existing
            // Excel spreadsheet
            
            Console.Write($"Writing data to Excel: {_excelFilename} ... ");

            ExcelWriter eWriter = new ExcelWriter(_excelFilename);
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

        /// <<summary>
        /// Command processor for extension
        /// looks for match between incoming command and the commands for the extension
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="parameters"></param>
        public override async Task<(bool, int)> ExecuteExtensionAsync(string cmd, string parameters)
        {
            bool matched = false;
            int result = 0;

            switch (cmd)
            {
                case "triptrackersync":
                case "tts":
                    result = await Initialize(parameters);
                    matched = true;
                    break;

                case "tdebug":
                    Debug = true;
                    matched = true;
                    break;

                case "tnodebug":
                    Debug = false;
                    matched = true;
                    break;
            }
            return (matched, result);
        }

        /// <summary>
        /// Print the help message for the Extension
        /// </summary>
        public override void Help()
        {
            Console.WriteLine(
                "\nExtension: Trip Tracker Sync - transfer events from device and write to an excel file\n" +
                "  triptrackersync <name>, <#>,\n" +
                "  or <address>\n" +
                "  F:=<path to excel file>),\n" +
                "  tts\t\t\t\t: connects to device,\n" +
                "  \t\t\t\t: sets the service to 'SimpleKeyService',\n" +
                "  \t\t\t\t: subscribes to characteristic #0,\n" +
                "  \t\t\t\t: gets the sync range,\n" +
                "  \t\t\t\t: transfers the events,\n" +
                "  \t\t\t\t: writes them to the excel file,\n" +
                "  \t\t\t\t: and updates the last sync Id on the device\n" +
                $" \t\t\t\t: F:= is optional and defaults to '{_excelFilename}\n" +
                "  tdebug\t\t\t: turns on debugging for the extension\n" +
                "  tnodebug\t\t\t: turns off debugging for the extension"
                );
        }

        /// <summary>
        /// Event handler for ValueChanged callback
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public override bool Characteristic_ValueChanged(GattCharacteristic sender, GattValueChangedEventArgs args)
        {
            CryptographicBuffer.CopyToByteArray(args.CharacteristicValue, out byte[] characteristicValue);
            bool processed;

            /*if (Debug)
            {
                Console.Write($"\n(tts) Value changed for {sender.Uuid} ({characteristicValue.Length} bytes):\n\thex:\t{BitConverter.ToString(characteristicValue).Replace("-", " ")}\n");
            }*/
            do
            {
                processed = false;
                string characteristicString = Encoding.UTF8.GetString(characteristicValue);

                if (characteristicString.StartsWith("SyncRange="))
                {
                    // Format is 'SyncRange=<StartId>,<EndId>\n'
                    var parts = characteristicString.Split('=');
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

                    parts = characteristicString.Split('\n');

                    if (Debug) Console.Write($"\t(tts) processing ({parts[0].Length + 1} bytes): {parts[0]} ==> ({StartId} - {EndId})\n");

                    // remove the part that has been processed
                    characteristicValue = characteristicValue.Skip(parts[0].Length + 1).ToArray();
                    processed = true;
                }
                else if (characteristicString.StartsWith("SyncData="))
                {
                    // The format of the command is SyncData=<hex Length of binary data>\n<start of binary data>
                    var parts = characteristicString.Split('=');
                    var parts2 = parts[1].Split('\n');
                    _aggregrateDataLengthRemaining = Convert.ToInt32(parts2[0], 16);

                    parts = characteristicString.Split('\n');
                    if (Debug) Console.Write($"\t(tts) processing ({parts[0].Length + 1}) {parts[0]} ==> {_aggregrateDataLengthRemaining}\n");

                    // remove the part that has been processed
                    characteristicValue = characteristicValue.Skip(parts[0].Length + 1).ToArray();

                    // setup for aggregation
                    _aggregrateDataInValueChanged = true;
                    _aggregateDataArray = new byte[0];
                    processed = true;
                }

                if (_aggregrateDataInValueChanged && (characteristicValue.Length > 0))
                {
                    // Aggregate data as long as there is data to aggregate. 
                    // There are two cases:
                    //    1) processing the rest of the data that came in the same packet as set up the aggregation
                    //    2) new packet arrived

                    if (characteristicValue.Length > _aggregrateDataLengthRemaining)
                    {
                        // there is more data than is needed in the characteristicValue, so just pull out what is needed and leave the rest
                        byte[] resultArray = new byte[_aggregateDataArray.Length + _aggregrateDataLengthRemaining];
                        Array.Copy(_aggregateDataArray, resultArray, _aggregateDataArray.Length);
                        Array.Copy(characteristicValue, 0, resultArray, _aggregateDataArray.Length + 1, _aggregrateDataLengthRemaining);
                        _aggregateDataArray = resultArray;

                        if (Debug) Console.Write($"\t(tts)*** Aggregating data (captured {_aggregrateDataLengthRemaining} bytes; 0 bytes remaining) ***\n");

                        characteristicValue = characteristicValue.Skip(_aggregrateDataLengthRemaining).ToArray();
                        _aggregrateDataLengthRemaining = 0;
                    }
                    else
                    {
                        // There is more data needed than is currently available so just copy it all into the aggregation array
                        _aggregateDataArray = _aggregateDataArray.Concat(characteristicValue).ToArray();
                        _aggregrateDataLengthRemaining -= characteristicValue.Length;
                        if (Debug) Console.Write($"\t(tts) *** Aggregating data (captured {characteristicValue.Length} bytes; {_aggregrateDataLengthRemaining} bytes remaining) ***\n");
                        characteristicValue = new byte[0];
                    }
                    processed = true;
                }

                if (_aggregrateDataInValueChanged && _aggregrateDataLengthRemaining == 0)
                {
                    // The aggregation has been completed, so now process the completed aggregate data.

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

                        case Event2Info.EventType.GeneratorOilChange:
                            var generatorOilChange = new GeneratorOilChangeInfo(_aggregateDataArray);
                            EventInfoList.Add(generatorOilChange);
                            generatorOilChange.Print();
                            break;


                        default:
                            Console.WriteLine($"ERROR: Aggregation of EventInfo got invalid type '{_aggregateDataArray[0]}");
                            break;
                    }
                    if (Debug)
                    {
                        Console.Write($"\t(tts) *** Aggregating data completed (processed {_aggregateDataArray.Length} bytes) {EventInfoList.Count} of {(EndId - StartId)} events captured ***\n");
                    }
                    _lastIdprocessed = EventInfoList.ElementAt(EventInfoList.Count - 1).Id;
                    // we have all the requested data 
                    // day summary is first, and then NumLegs 
                    // and we have all the Event info
                    if (EventInfoList.Count >= (EndId - StartId))
                    {
                        PendingWork = PendingWorkType.ProcessResults;
                        if (Debug) Console.Write($"\t(tts) *** All requested Events {(EndId - StartId)} have been received ***\n");
                    }
                }
            } while (processed && characteristicValue.Length > 0);

            return processed;
        }
    }
}
