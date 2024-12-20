using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Security.Cryptography;

namespace BLEConsole
{
    internal class GPXSync : ExtensionBase
    {
        internal enum PendingWorkType { None, AckChunk, CloseFile, Done };
        internal enum AggregationType { None, Command, Data };
        internal enum CommandType { None, GPXRange, GPXData };

        internal PendingWorkType PendingWork { get; set; } = PendingWorkType.None;

        string _gpxPath = "C:\\Users\\drmcl\\GitHub\\Temp";
        string _gpxFilenameTemplate = "Trip{0}.gpx";
        bool _haveStartEnd = false;

        AggregationType _aggregationType = AggregationType.None;
        CommandType _commandType = CommandType.None;

        int _currentChunkNum;
        int _totalChunks;
        int _bytesInChunkRemaining;
 
        byte[] _aggregateDataArray = null;
        FileWriter _gpxFileWriter = null;

        int StartId { get; set; } = 0;
        int EndId { get; set; } = 0;

        public bool Debug { get; set; } = false;

        private (string filename, string parameters) ExtractPath(string parameters)
        {
            string pattern = @"P:=""([^""]+)""|F:=([^ ]+)";
            var match = Regex.Match(parameters, pattern);

            if (match.Success)
            {
                string filename = match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value;
                string modifiedString = parameters.Remove(match.Index, match.Length).Trim();
                return (filename, modifiedString);
            }

            return (null, parameters);
        }

        public async Task<int> Initialize(string input)
        {
            var (path, parameters) = ExtractPath(input);
            if (path != null) _gpxPath = path;
            Console.WriteLine($"Will write the gpx files to the folder {_gpxPath} ");

            if (!Directory.Exists(_gpxPath))
            {
                Console.WriteLine($"Aborting: {_gpxPath} does not exist");
                return 1;
            }

            int result = 0;
            try
            {
                result += await Program.OpenDevice(parameters);

                if (result == 0)
                    result += await Program.SetService("SimpleKeyService");

                if (result == 0)
                    result += await Program.SubscribeToCharacteristic("#0"); // SimpleKeyState

                // request starting and ending Ids of the gpx files to sync
                // result will be processed in the Characteristic_ValueChanged Handler
                if (result == 0)
                    result += await Program.WriteCharacteristic("#0 XC?");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                result++;
            }
            return result;
        }

        public async Task<int> TransferFile (string input)
        {
            int result = 0;
            if (!string.IsNullOrEmpty(input)) 
            {
                try
                {
                    var fileNumber = Convert.ToInt32(input);

                    if (_gpxFileWriter == null) _gpxFileWriter = new FileWriter();
                    string filename = _gpxFileWriter.CreateFile(_gpxPath, _gpxFilenameTemplate, fileNumber);
                    Console.WriteLine($"Writing to {filename}");

                    string cmd = $"#0 XD?{fileNumber};";
                    if (Debug) Console.WriteLine($"\nGPX File request {fileNumber} {cmd}");
                    result += await Program.WriteCharacteristic(cmd);

                    if (result == 0)
                    {
                        PendingWork = PendingWorkType.None;
                        result += await ProcessPendingWork();
                    }
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"Request GPX File Transfer: Exception {ex}");
                }
            }
            else
            {
                Console.WriteLine("fileNumber can not be empty.");
                result += 1;
            }
            if (result > 0) Console.WriteLine("Request GPX File Transfer failed");

            return result;
        }

        /// <summary>
        /// Async method to process work generated in Characteristic_ValueChanged
        /// </summary>
        private async Task<int> ProcessPendingWork()
        {
            int result = 0;

            while (PendingWork != PendingWorkType.Done)
            {
                switch (PendingWork)
                {
                    case PendingWorkType.AckChunk:
                        PendingWork = PendingWorkType.None;
                        string cmd = $"#0 XA={_currentChunkNum};";
                        if (Debug) 
                            Console.WriteLine($"\nGPX Ack Chunk {_currentChunkNum} {cmd}");
                        else
                        {
                            if (_currentChunkNum % 80 == 0)
                                Console.WriteLine($". ({(double)_currentChunkNum / _totalChunks * 100:F1}%)");
                            else
                                Console.Write(".");
                        }
                        result += await Program.WriteCharacteristic(cmd);
                        break;

                    case PendingWorkType.CloseFile:
                        PendingWork = PendingWorkType.Done;
                        _gpxFileWriter.CloseFile();
                        if (!Debug) Console.WriteLine("<done>");
                        break;

                    default:
                        Thread.Sleep(200);
                        break;
                }
                if (result > 0) break;
            }
            return result ;
        }

        /// <summary>
        /// Extract GPXRange parameters
        /// </summary>
        /// <param name="data"></param>
        private bool CaptureGPXRange(string data)
        {
            var counts = data.Split(',');
            if (counts.Length == 2)
            {
                StartId = Convert.ToInt32(counts[0], 10);
                EndId = Convert.ToInt32(counts[1], 10);
                _haveStartEnd = EndId > StartId;
            }
            else
                _haveStartEnd = false;

            return _haveStartEnd;
        }

        /// <summary>
        /// Extract GPXData parameters
        /// </summary>
        /// <param name="data"></param>
        private bool CaptureGPXData(string data)
        {
            var counts = data.Split(',');
            if (counts.Length == 3)
            {
                _currentChunkNum = Convert.ToInt32(counts[0], 10);
                _totalChunks = Convert.ToInt32(counts[1], 10);
                _bytesInChunkRemaining = Convert.ToInt32(counts[2], 16);
                if (_totalChunks > 0) return true;
                else return false;
            }
            else
                return false; 
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
                case "gpxsyncinit":
                case "gsi":
                    result = await Initialize(parameters);
                    matched = true;
                    break;

                case "gpxsyncfile":
                case "gsf":
                    result = await TransferFile(parameters);
                    matched = true;
                    break;

                case "gdebug":
                    Debug = true;
                    matched = true;
                    break;

                case "gnodebug":
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
                "\nExtension: GPXSync - transfer GPX files from device\n" +
                "  gpxsyncinit <name>, <#>,\n" +
                "  or <address>\n" +
                "  P:=<path to write GPX files>),\n" +
                "  gsi\t\t\t\t: connects to device,\n" +
                "  \t\t\t\t: sets the service to 'SimpleKeyService',\n" +
                "  \t\t\t\t: subscribes to characteristic #0,\n" +
                "  \t\t\t\t: and gets the sync range\n" +
                $" \t\t\t\t: P:= is optional and defaults to '{_gpxPath}\n" +
                "  gpxsyncfile, gsf <id>" +
                $"  \t: transfers the <id> gpx file from the device and writes it to {_gpxFilenameTemplate}\n" +
                "  gdebug\t\t\t: turns on debugging for the extension\n" +
                "  gnodebug\t\t\t: turns off debugging for the extension"
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

            if (Debug) Console.Write($"\n(gpx) Value changed for {sender.Uuid} ({characteristicValue.Length} bytes):\n");

            do
            {
                processed = false;
                string characteristicString = Encoding.UTF8.GetString(characteristicValue);

                switch (_aggregationType)
                {
                    case AggregationType.None:
                        if (characteristicString.StartsWith("GPXRange="))
                        {
                            // Format is 'GPXRange=<StartId>,<EndId>\n'
                            // parameters will be extracted in AggregationType.Command section
                            // parameters could be longer than what is left in the current characteristicString.
                            var parts = characteristicString.Split('=');

                            if (Debug) Console.Write($"\t(gpx) processing ({parts[0].Length + 1} bytes): {parts[0]} -- remaining ({parts[1].Length})\n");

                            // remove the part that has been processed
                            characteristicValue = characteristicValue.Skip(parts[0].Length + 1).ToArray();

                            _aggregationType = AggregationType.Command;
                            _commandType = CommandType.GPXRange; // set up to process the parameters
                            _aggregateDataArray = new byte[0];
                            processed = true;
                        }

                        if (characteristicString.StartsWith("GPXData="))
                        {
                            // The format of the command id GPXData=<current Chunk#>,<total Chunks>,<hex Length of chunk>\n<start of binary data>
                            // parameters will be extracted in Aggregation.Command section
                            // parameters could be longer than what is left in the current characteristicString.
                            var parts = characteristicString.Split('=');

                            if (Debug) Console.Write($"\t(gpx) processing ({parts[0].Length + 1} bytes): {parts[0]} -- remaining ({parts[1].Length})\n");

                            // remove the part that has been processed
                            characteristicValue = characteristicValue.Skip(parts[0].Length + 1).ToArray();

                            _aggregationType = AggregationType.Command;
                            _commandType = CommandType.GPXData;
                            _aggregateDataArray = new byte[0];
                            processed = true;
                        }

                        if (characteristicString.StartsWith("GPXDone"))
                        {
                            // File has been transferred completely.
                            var parts = characteristicString.Split('\n');
                            if (Debug) Console.Write($"\t(gpx) processing ({parts[0].Length + 1} bytes): {parts[0]} -- remaining ({parts[1].Length})\n");

                            // remove the part that has been processed
                            characteristicValue = characteristicValue.Skip(parts[0].Length + 1).ToArray();
                            processed = true;
                            PendingWork = PendingWorkType.CloseFile;
                        }
                        break;

                    case AggregationType.Command:
                        int newlineIndex = Array.IndexOf(characteristicValue, (byte)'\n');
                        if (newlineIndex >= 0)
                        {
                            // \n is not the first character
                            if (newlineIndex > 0)
                            {
                                if (_aggregateDataArray.Length > 0)
                                {
                                    // There is already data in the aggregateDataArray
                                    byte[] resultArray = new byte[_aggregateDataArray.Length + newlineIndex];
                                    Array.Copy(_aggregateDataArray, resultArray, _aggregateDataArray.Length);
                                    Array.Copy(characteristicValue, 0, resultArray, _aggregateDataArray.Length, newlineIndex);
                                    _aggregateDataArray = resultArray;
                                }
                                else
                                {
                                    _aggregateDataArray = new byte[newlineIndex];
                                    Array.Copy(characteristicValue, _aggregateDataArray, newlineIndex);
                                }
                            }

                            if (Debug) Console.Write($"\t(gpx)*** Aggregating command (captured {newlineIndex + 1} bytes) ***\n");

                            // remove the command data from the characteristicValue.
                            // This is the data that still needs to be processed.
                            characteristicValue = characteristicValue.Skip(newlineIndex + 1).ToArray();

                            string dataArray = Encoding.UTF8.GetString(_aggregateDataArray);

                            switch (_commandType)
                            {
                            case CommandType.GPXData:
                                if (CaptureGPXData(dataArray))
                                {
                                    if (Debug) Console.Write($"\t(gpx) processing ({dataArray.Length + 1}) GPXData {dataArray} ==> {_bytesInChunkRemaining}\n");
                                    _aggregationType = AggregationType.Data;
                                    _aggregateDataArray = new byte[0];
                                }
                                else
                                {
                                    Console.WriteLine($"Error: Failed to get parameters for GPXData ({dataArray})");
                                    _aggregationType = AggregationType.None;
                                    PendingWork = PendingWorkType.Done;
                                }
                                break;

                            case CommandType.GPXRange:
                                if (CaptureGPXRange(dataArray))
                                {
                                    if (Debug) Console.WriteLine($"\t(gpx) processing ({dataArray.Length + 1} bytes): GPXRange {dataArray} ==> ({StartId} - {EndId})");
                                    else Console.WriteLine($"\tGPX Range ==> ({StartId} - {EndId})");
                                    _aggregationType = AggregationType.None;
                                }
                                else
                                {
                                    Console.WriteLine($"Error: Failed to get parameters for GPXRange ({dataArray})");
                                        _aggregationType = AggregationType.None;
                                }
                                break;
                            }
                            _commandType = CommandType.None;
                        }
                        else
                        {
                            // There is more command data needed than is currently available so just copy it all into the aggregation array
                            _aggregateDataArray = _aggregateDataArray.Concat(characteristicValue).ToArray();
                            if (Debug) Console.Write($"\t(gpx) *** Aggregating command (captured {characteristicValue.Length} bytes ***\n");
                            characteristicValue = new byte[0];
                        }
                        processed = true;
                        break;

                    case AggregationType.Data:
                        // Aggregate data as long as there is data to aggregate. 
                        // There are two cases:
                        //    1) processing the rest of the data that came in the same packet as set up the aggregation
                        //    2) new packet arrived

                        if (characteristicValue.Length >= _bytesInChunkRemaining)
                        {
                            // there is more data than is needed in the characteristicValue, so just pull out what is needed and leave the rest
                            byte[] resultArray = new byte[_aggregateDataArray.Length + _bytesInChunkRemaining];
                            Array.Copy(_aggregateDataArray, resultArray, _aggregateDataArray.Length);
                            Array.Copy(characteristicValue, 0, resultArray, _aggregateDataArray.Length, _bytesInChunkRemaining);
                            _aggregateDataArray = resultArray;

                            characteristicValue = characteristicValue.Skip(_bytesInChunkRemaining).ToArray();
                            if (Debug) Console.Write($"\t(gpx)*** Aggregating data (captured {_bytesInChunkRemaining} bytes; 0 bytes needed) -- remaining ({characteristicValue.Length}) ***\n");

                            _bytesInChunkRemaining = 0;

                            // The aggregation has been completed, so now write the aggregated data array to the file.
                            _gpxFileWriter.WriteData(_aggregateDataArray);
                            PendingWork = PendingWorkType.AckChunk;
                            _aggregationType = AggregationType.None;

                            if (Debug) Console.Write($"\t(gpx) *** Aggregating data completed (processed {_aggregateDataArray.Length} bytes) ***\n");
                        }
                        else
                        {
                            // There is more data needed than is currently available so just copy it all into the aggregation array
                            _aggregateDataArray = _aggregateDataArray.Concat(characteristicValue).ToArray();
                            _bytesInChunkRemaining -= characteristicValue.Length;
                            if (Debug) Console.Write($"\t(gpx) *** Aggregating data (captured {characteristicValue.Length} bytes; {_bytesInChunkRemaining} bytes needed) ***\n");
                            characteristicValue = new byte[0];
                        }
                        processed = true;

                        break;
                }
            } while (processed && characteristicValue.Length > 0);

            return processed;
        }
    }
}
