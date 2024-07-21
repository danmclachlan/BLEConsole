using Microsoft.Office.Interop.Excel;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace BLEConsole
{
    public static class TripTracker
    {
        enum AggregateDataType { None, VehicleInfo, TripInfo };
        enum PendingWorkType { None, RequestData, ProcessResults, Done };

        static bool _aggregrateDataInValueChanged = false;
        static int _aggregrateDataLengthRemaining = 0;
        static byte[] _aggregateDataArray = null;
        static AggregateDataType _aggregateDataType = AggregateDataType.None;
        static PendingWorkType PendingWork { get; set; } = PendingWorkType.None;
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

                // start the listener for processing the results
                PendingWork = PendingWorkType.None;

                if (result == 0)
                {
                    // request the number of Legs in the current Day
                    // result will be processed in the Characteristic_ValueChanged Handler
                    result += await Program.WriteCharacteristic("#0 LC?");
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

        enum ExcelInsertType { DayStart, LegStart, LegEnd, DayEnd };

        public static void ProcessTripData()
        {
            Console.WriteLine("Vehicle Info");
            VehicleInfo.Print();

            for (int i = 0; i < TripInfoList.Count; i++)
            {
                if (i == 0)
                {
                    Console.WriteLine($"Day Info"); ;
                }
                else
                {
                    Console.WriteLine($"Trip Leg {i} Info");
                }
                TripInfoList[i].Print();
            }

            Application excelApp = new Excel.Application();
            // Make the object visible.
            //excelApp.Visible = true;
            var filename = "C:\\Users\\drmcl\\GitHub\\Temp\\CA-2024-07-Trip.xlsx";

            Console.Write($"Writing data to Excel: {filename} ... ");

            Workbook workbook = excelApp.Workbooks.Open("C:\\Users\\drmcl\\GitHub\\Temp\\CA-2024-07-Trip.xlsx");
            Worksheet worksheet = workbook.Sheets[1];

            ListObject table = worksheet.ListObjects["TripDetail"];

            for (int i = 0; i < TripInfoList.Count; i++)
            {
                if (i == 0)
                {
                    InsertRowIntoTable(table, VehicleInfo.Name, "Focus", ExcelInsertType.DayStart, TripInfoList[i]);
                }
                else
                {
                    InsertRowIntoTable(table, VehicleInfo.Name, "Focus", ExcelInsertType.LegStart, TripInfoList[i]);
                    InsertRowIntoTable(table, VehicleInfo.Name, "Focus", ExcelInsertType.LegEnd, TripInfoList[i]);
                }
            }

            InsertRowIntoTable(table, VehicleInfo.Name, "Focus", ExcelInsertType.DayEnd, TripInfoList[0]);

            workbook.Save();
            workbook.Close();
            excelApp.Quit();
            Console.WriteLine("Complete");
        }

        // DateTime	Vehicle	Tow	Type	Description	gallons	Price	Odometer	Engine Hrs Counter	Fuel Level	Leg Duration	Day Duration	Distance Traveled	Engine Hrs Used	Fuel Used

        static TimeSpan LegDurationSum { get; set; } = new TimeSpan(0);

        static void InsertRowIntoTable(ListObject table,
            string vehicle, string towVehicle,
            ExcelInsertType insertType,
            TripInfo trip)
        {
            TimeSpan legDuration;

            ListRow newRow = table.ListRows.Add();

            newRow.Range[1, 2].Value = vehicle;
            newRow.Range[1, 3].Value = towVehicle;

            switch (insertType)
            {
                case ExcelInsertType.DayStart:
                    newRow.Range[1, 4].Value = "Start";
                    LegDurationSum = new TimeSpan(0);
                    break;
                case ExcelInsertType.DayEnd:
                    newRow.Range[1, 4].Value = "End";
                    break;
                case ExcelInsertType.LegStart:
                    newRow.Range[1, 4].Value = "Depart";
                    break;
                case ExcelInsertType.LegEnd:
                    newRow.Range[1, 4].Value = "Arrive";
                    break;
            }

            switch (insertType)
            {
                case ExcelInsertType.DayStart:
                case ExcelInsertType.LegStart:
                    newRow.Range[1, 1].Value = trip.StartLocalTime;
                    newRow.Range[1, 8].Value = trip.StartOdometer;
                    newRow.Range[1, 9].Value = trip.StartEngineHours;
                    newRow.Range[1,10].Value = trip.StartFuel;
                    break;
                case ExcelInsertType.DayEnd:
                case ExcelInsertType.LegEnd:
                    newRow.Range[1, 1].Value = trip.EndLocalTime;
                    newRow.Range[1, 8].Value = trip.EndOdometer;
                    newRow.Range[1, 9].Value = trip.EndEngineHours;
                    newRow.Range[1,10].Value = trip.EndFuel;

                    legDuration = trip.EndTimeGMT - trip.StartTimeGMT;
                    if (insertType == ExcelInsertType.LegEnd)
                    {
                        newRow.Range[1,11].Value = legDuration.TotalSeconds / 86400; // Excel stores time as a fraction of a day
                        LegDurationSum += legDuration;
                    }
                    else
                    {
                        newRow.Range[1,11].Value = LegDurationSum.TotalSeconds / 86400; // Excel stores time as a fraction of a day
                        newRow.Range[1,12].Value = legDuration.TotalSeconds / 86400; // Excel stores time as a fraction of a day
                    }
                    newRow.Range[1,13].Value = trip.EndOdometer - trip.StartOdometer;
                    newRow.Range[1,14].Value = trip.EndEngineHours - trip.StartEngineHours;
                    newRow.Range[1,15].Value = trip.FuelUsed;
                    if (trip.FuelUsed > 0)
                        newRow.Range[1, 16].Value = (trip.EndOdometer - trip.StartOdometer) / trip.FuelUsed;

                    break;
            }
            for (int i = 6; i <= 16; i++)
            {
                if (i == 6)
                    newRow.Range[1, i].NumberFormat = "0.000";
                else if (i == 7)
                    newRow.Range[1, i].NumberFormat = "$0.00";
                if (i == 11 || i == 12)
                    newRow.Range[1, i].NumberFormat = "[h]:mm:ss";
                else 
                    newRow.Range[1, i].NumberFormat = "0.0";
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
                NumLegs = Convert.ToInt32(parts2[0], 10);
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
                    }
                    else if (_aggregateDataType == AggregateDataType.TripInfo)
                    {
                        var info = new TripInfo(_aggregateDataArray);
                        TripInfoList.Add(info);

                        // we have all the requested data 
                        // day summary is first, and then NumLegs 
                        if (TripInfoList.Count >= NumLegs + 1)
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
        public int StartTimeTZOffset { get; set; }
        public DateTime StartTimeGMT { get { return new DateTime(1970, 1, 1).AddSeconds(StartTime); } }
        public DateTime StartLocalTime { get { return StartTimeGMT.AddHours(StartTimeTZOffset); } }
        public DateTime EndTimeGMT { get {  return new DateTime(1970, 1, 1).AddSeconds(EndTime); } }
        public DateTime EndLocalTime { get { return EndTimeGMT.AddHours(EndTimeTZOffset); } }
        public uint EndTime { get; set; }
        public int EndTimeTZOffset { get; set; }
        public double StartOdometer { get; set; }
        public double EndOdometer { get; set; }
        public double StartEngineHours { get; set; }
        public double EndEngineHours { get; set; }
        public double StartFuel { get; set; }
        public double EndFuel { get; set; }
        public double FuelUsed {  get; set; }

        public TripInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                Type = new string(reader.ReadChars(1));
                Id = reader.ReadUInt32();
                StartTime = reader.ReadUInt32();
                EndTime = reader.ReadUInt32();

                StartTimeTZOffset = reader.ReadInt32();
                EndTimeTZOffset = reader.ReadInt32();

                StartOdometer = reader.ReadDouble();
                EndOdometer = reader.ReadDouble();
                StartEngineHours = reader.ReadDouble();
                EndEngineHours = reader.ReadDouble();
                StartFuel = reader.ReadDouble();
                EndFuel = reader.ReadDouble();
                FuelUsed = reader.ReadDouble();
            }
        }

        public void Print()
        {
            //var PDToffset = -7;  // UTC -7
            //DateTime start = new DateTime(1970, 1, 1).AddSeconds(StartTime);
            //DateTime startPDT = start.AddHours(PDToffset);
            //DateTime startLocal = start.ToLocalTime();
            //DateTime end = new DateTime(1970, 1, 1).AddSeconds(EndTime);
            //DateTime endPDT = end.AddHours(PDToffset);
            //DateTime endLocal = end.ToLocalTime();

            Console.WriteLine($"\tStart Time: {StartTime} ({StartLocalTime}) GMT{StartTimeTZOffset}");
            Console.WriteLine($"\tEnd Time: {EndTime} ({EndLocalTime}) GMT{EndTimeTZOffset}");
            Console.WriteLine($"\tStart Odometer: {StartOdometer:F1}");
            Console.WriteLine($"\tEnd Odometer: {EndOdometer:F1}");
            Console.WriteLine($"\tStart Engine Hours: {StartEngineHours:F1}");
            Console.WriteLine($"\tEnd Engine Hours: {EndEngineHours:F1}");
            Console.WriteLine($"\tStart Fuel: {StartFuel:F1}");
            Console.WriteLine($"\tEnd Fuel: {EndFuel:F1}");
            Console.WriteLine($"\tFuel Used: {FuelUsed:F1}");
        }
    }

}
