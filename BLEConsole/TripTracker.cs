using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Printers;
using Windows.Security.Cryptography;

namespace BLEConsole
{
    public static class TripTracker
    {
        enum AggregateDataType { None, VehicleInfo, TripInfo };

        static bool _aggregrateDataInValueChanged = false;
        static int _aggregrateDataLengthRemaining = 0;
        static byte[] _aggregateDataArray = null;
        static AggregateDataType _aggregateDataType = AggregateDataType.None;
        public static bool PendingWork { get; set; } = false;
        static int NumLegs { get; set; } = 0;
        public static bool Debug { get; set; } = false;

        public static List<TripInfo> TripInfoList { get; set; } = new List<TripInfo>();
        public static VehicleInfo VehicleInfo { get; set; }

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

                if ( result == 0)
                {
                    // request the number of Legs in the current Day
                    // result will be processed in the Characteristic_ValueChanged Handler
                    result += await Program.WriteCharacteristic("#0 LC?");
                    Thread.Sleep(200);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                result++;
            }
            return result;
        }

        public static async Task RequestTripData()
        {
            int result = 0;

            PendingWork = false;

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

            } catch (Exception ex)
            {
                Console.WriteLine($"RequestTripData: Exception {ex}");
                result++;
            }
            if (result > 0)
            {
                Console.WriteLine($"RequestTripData failed");
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

            if (tempValue.StartsWith("LegCount="))
            {
                var parts = tempValue.Split('=');
                var parts2 = parts[1].Split('\n');
                NumLegs = Convert.ToInt32(parts2[0],10);
                PendingWork = true;
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

            if (startAggregation)
            {
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

                        Console.WriteLine("Trip Leg Info");
                        info.Print();
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

    public class VehicleInfo
    {
        public string Name { get; set; }
        public double Odometer { get; set; }
        public double EngineHours { get; set; }
        public double OdometerBase { get; set; }
        public double EngineHoursBase { get; set; }
        public double FuelCapacity { get; set; }
        public double FuelReserve { get; set; }
        public double FuelFillUpMileage { get; set; }
        public double OilChangeInterval { get; set; }
        public double OilChangeMileage { get; set; }

        public VehicleInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                this.Name = new string(reader.ReadChars(10));
                this.Odometer = reader.ReadDouble();
                this.EngineHours = reader.ReadDouble();
                this.OdometerBase = reader.ReadDouble();
                this.EngineHoursBase = reader.ReadDouble();
                this.FuelCapacity = reader.ReadDouble();
                this.FuelReserve = reader.ReadDouble();
                this.FuelFillUpMileage = reader.ReadDouble();
                this.OilChangeInterval = reader.ReadDouble();
                this.OilChangeMileage = reader.ReadDouble();
            }
        }

        public void Print()
        {
            Console.Write($"\tName: {Name}\n");
            Console.Write($"\tOdometer: {Odometer:F1}\n");
            Console.Write($"\tEngine Hours: {EngineHours:F1}\n");
            Console.Write($"\tOdometer Base: {OdometerBase:F1}\n");
            Console.Write($"\tEngine Hours Base: {EngineHoursBase:F1}\n");
            Console.Write($"\tFuel Capacity: {FuelCapacity:F1}\n");
            Console.Write($"\tFuel Reserve: {FuelReserve:F1}\n");
            Console.Write($"\tFuel Fillup Mileage: {FuelFillUpMileage:F1}\n");
            Console.Write($"\tOil Change Interval: {OilChangeInterval:F1}\n");
            Console.Write($"\tOil Change Mileage: {OilChangeMileage:F1}\n");
        }
    }

    public class TripInfo
    {
        public string Type { get; set; }
        public uint Id { get; set; }
        public uint StartTime { get; set; }
        public uint EndTime { get; set; }
        public double StartOdometer { get; set; }
        public double EndOdometer { get; set; }
        public double StartEngineHours { get; set; }
        public double EndEngineHours { get; set; }
        public double StartFuel { get; set; }
        public double EndFuel { get; set; }

        public TripInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                Type = new string(reader.ReadChars(1));
                Id = reader.ReadUInt32();
                StartTime = reader.ReadUInt32();
                EndTime = reader.ReadUInt32();
                StartOdometer = reader.ReadDouble();
                EndOdometer = reader.ReadDouble();
                StartEngineHours = reader.ReadDouble();
                EndEngineHours = reader.ReadDouble();
                StartFuel = reader.ReadDouble();
                EndFuel = reader.ReadDouble();
            }
        }

        public void Print()
        {
            DateTime start = new DateTime(1970, 1, 1).ToLocalTime().AddSeconds(StartTime);
            DateTime end = new DateTime(1970, 1, 1).ToLocalTime().AddSeconds(EndTime);

            Console.WriteLine($"\tStart Time: {StartTime} ({start})");
            Console.WriteLine($"\tEnd Time: {EndTime} ({end}");
            Console.WriteLine($"\tStart Odometer: {StartOdometer:F1}");
            Console.WriteLine($"\tEnd Odometer: {EndOdometer:F1}");
            Console.WriteLine($"\tStart Engine Hours: {StartEngineHours:F1}");
            Console.WriteLine($"\tEnd Engine Hours: {EndEngineHours:F1}");
            Console.WriteLine($"\tStart Fuel: {StartFuel:F1}");
            Console.WriteLine($"\tEnd Fuel: {EndFuel:F1}");
        }
    }

}
