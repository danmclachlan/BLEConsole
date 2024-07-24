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
                        if(Debug) Console.WriteLine($"\nEvent request {i} {cmd}");
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
            // Print out on the console all the trip data collected from the TripTracker
            // system.
            Console.WriteLine("Vehicle Info");
            VehicleInfo.Print();

            for (int i = 0; i < TripInfoList.Count; i++)
            {
                if (i == 0)
                    Console.WriteLine($"Day Info");
                else
                    Console.WriteLine($"Trip Leg {i} Info");

                TripInfoList[i].Print();
            }

            for (int i = 0; i < EventInfoList.Count; i++)
            {
                if (EventInfoList[i] is PurchaseFuelInfo purchaseFuel)
                    purchaseFuel.Print();
                else if (EventInfoList[i] is PurchasePropaneInfo purchasePropane)
                    purchasePropane.Print();
                else if (EventInfoList[i] is ChangeOilInfo changeOil)
                    changeOil.Print();
                else
                    EventInfoList[i].Print();
            }

            // Store all the Trip Tracker data for the day into an existing
            // Excel spreadsheet
            // TODO: make the spreadsheet be setable rather than a constant.

            Application excelApp = new Excel.Application();
            // Make the object visible.
            //excelApp.Visible = true;
            var filename = "C:\\Users\\drmcl\\GitHub\\Temp\\CA-2024-07-Trip.xlsx";

            Console.Write($"Writing data to Excel: {filename} ... ");

            Workbook workbook = excelApp.Workbooks.Open("C:\\Users\\drmcl\\GitHub\\Temp\\CA-2024-07-Trip.xlsx");
            Worksheet worksheet = workbook.Sheets[1];

            ListObject table = worksheet.ListObjects["TripDetail"];

            // Insert individual lines in the TripDetail table, one line for each Event.
            // An event can be one of the following:
            // - Start of the Day 
            // - Start of a Leg
            // - End of a Leg
            // - Purchase Fuel
            // - Purchase Propane
            // - Change Oil
            // - End of the Day
            // Events are written in time assending order.

            int j = 0;
            for (int i = 0; i < TripInfoList.Count; i++)
            {
                // Find any events that occurred before the start of this Day/Leg
                if (j < EventInfoList.Count && EventInfoList[j].Time < TripInfoList[i].StartTime)
                    InsertEventRowIntoTable(table, EventInfoList[j++]);

                if (i == 0)
                    InsertTripRowIntoTable(table, VehicleInfo.Name, "Focus", ExcelInsertType.DayStart, TripInfoList[i]);
                else
                {
                    InsertTripRowIntoTable(table, VehicleInfo.Name, "Focus", ExcelInsertType.LegStart, TripInfoList[i]);
                    InsertTripRowIntoTable(table, VehicleInfo.Name, "Focus", ExcelInsertType.LegEnd, TripInfoList[i]);
                }
            }
            // Find any remaining event that occurred after the end of the last leg
            while (j < EventInfoList.Count)
                InsertEventRowIntoTable(table, EventInfoList[j++]);

            InsertTripRowIntoTable(table, VehicleInfo.Name, "Focus", ExcelInsertType.DayEnd, TripInfoList[0]);

            workbook.Save();
            workbook.Close();
            excelApp.Quit();
            Console.WriteLine("Complete");
        }

        // Definitions of the Columns in the "TripDetail" table in the Excel worksheet
        public static class ExcelRow
        {
            public const int LocalDateTime = 1;
            public const int TimezoneOffset = 2;
            public const int Vehicle = 3;
            public const int TowVehicle = 4;
            public const int Type = 5;
            public const int GPSLocation = 6;
            public const int Description = 7;
            public const int Quantity = 8;
            public const int Cost = 9;
            public const int Odometer = 10;
            public const int EngineHoursCounter = 11;
            public const int FuelLevel = 12;
            public const int LegDuration = 13;
            public const int DayDuration = 14;
            public const int DistanceTraveled = 15;
            public const int EngineHoursUsed = 16;
            public const int FuelUsed = 17;
            public const int MPG = 18;
            public const int MaxRows = 18;
        }

        // Need to sum the durations of each of the legs to be able to provide
        // actual traveling time in addition to to elapsed time for the day.
        static TimeSpan LegDurationSum { get; set; } = new TimeSpan(0);

        static void InsertTripRowIntoTable(ListObject table,
            string vehicle, string towVehicle,
            ExcelInsertType insertType,
            TripInfo trip)
        {
            TimeSpan legDuration;

            ListRow newRow = table.ListRows.Add();

            newRow.Range[1, ExcelRow.Vehicle].Value = vehicle;
            newRow.Range[1, ExcelRow.TowVehicle].Value = towVehicle;

            switch (insertType)
            {
                case ExcelInsertType.DayStart:
                    newRow.Range[1, ExcelRow.Type].Value = "Start";
                    LegDurationSum = new TimeSpan(0);
                    break;
                case ExcelInsertType.DayEnd:
                    newRow.Range[1, ExcelRow.Type].Value = "End";
                    break;
                case ExcelInsertType.LegStart:
                    newRow.Range[1, ExcelRow.Type].Value = "Depart";
                    break;
                case ExcelInsertType.LegEnd:
                    newRow.Range[1, ExcelRow.Type].Value = "Arrive";
                    break;
            }

            switch (insertType)
            {
                case ExcelInsertType.DayStart:
                case ExcelInsertType.LegStart:
                    newRow.Range[1, ExcelRow.LocalDateTime].Value = trip.StartLocalTime;
                    newRow.Range[1, ExcelRow.TimezoneOffset].Value = trip.StartTimeTZOffset;
                    if (trip.StartGPSFixValid)
                    {
                        string display = $"({trip.StartLatitude:F7}, {trip.StartLongitude:F7})";
                        string googleMapsUrl = $"https://www.google.com/maps?q={trip.StartLatitude},{trip.StartLongitude}";

                        // Get the worksheet from the table
                        Worksheet parentWorksheet = (Worksheet)table.Parent;
                        parentWorksheet.Hyperlinks.Add(newRow.Range[1, ExcelRow.GPSLocation], googleMapsUrl, Type.Missing, "Open Google Maps", display);
                    }
                    newRow.Range[1,ExcelRow.Odometer].Value = trip.StartOdometer;
                    newRow.Range[1,ExcelRow.EngineHoursCounter].Value = trip.StartEngineHours;
                    newRow.Range[1,ExcelRow.FuelLevel].Value = trip.StartFuel;
                    break;
                case ExcelInsertType.DayEnd:
                case ExcelInsertType.LegEnd:
                    newRow.Range[1, ExcelRow.LocalDateTime].Value = trip.EndLocalTime;
                    newRow.Range[1, ExcelRow.TimezoneOffset].Value = trip.EndTimeTZOffset;
                    if (trip.EndGPSFixValid)
                    {
                        string display = $"({trip.EndLatitude:F7}, {trip.EndLongitude:F7})";
                        string googleMapsUrl = $"https://www.google.com/maps?q={trip.EndLatitude},{trip.EndLongitude}";

                        // Get the worksheet from the table
                        Worksheet parentWorksheet = (Worksheet)table.Parent;
                        parentWorksheet.Hyperlinks.Add(newRow.Range[1, ExcelRow.GPSLocation], googleMapsUrl, Type.Missing, "Open Google Maps", display);
                    }
                    newRow.Range[1, ExcelRow.Odometer].Value = trip.EndOdometer;
                    newRow.Range[1, ExcelRow.EngineHoursCounter].Value = trip.EndEngineHours;
                    newRow.Range[1, ExcelRow.FuelLevel].Value = trip.EndFuel;

                    legDuration = trip.EndTimeGMT - trip.StartTimeGMT;
                    if (insertType == ExcelInsertType.LegEnd)
                    {
                        newRow.Range[1,ExcelRow.LegDuration].Value = legDuration.TotalSeconds / 86400; // Excel stores time as a fraction of a day
                        LegDurationSum += legDuration;
                    }
                    else
                    {
                        newRow.Range[1,ExcelRow.LegDuration].Value = LegDurationSum.TotalSeconds / 86400; // Excel stores time as a fraction of a day
                        newRow.Range[1,ExcelRow.DayDuration].Value = legDuration.TotalSeconds / 86400; // Excel stores time as a fraction of a day
                    }
                    newRow.Range[1,ExcelRow.DistanceTraveled].Value = trip.EndOdometer - trip.StartOdometer;
                    newRow.Range[1,ExcelRow.EngineHoursUsed].Value = trip.EndEngineHours - trip.StartEngineHours;
                    newRow.Range[1,ExcelRow.FuelUsed].Value = trip.FuelUsed;
                    if (trip.FuelUsed > 0)
                        newRow.Range[1, ExcelRow.MPG].Value = (trip.EndOdometer - trip.StartOdometer) / trip.FuelUsed;

                    break;
            }

            // Format the cells in the added row to ensure they are formatted as expected.
            for (int i = 1; i <= ExcelRow.MaxRows; i++)
            {
                switch (i)
                {
                    case ExcelRow.TimezoneOffset:
                        newRow.Range[1, i].NumberFormat = "0";
                        break;
                    case ExcelRow.Quantity:
                        newRow.Range[1, i].NumberFormat = "0.000";
                        break;
                    case ExcelRow.Cost:
                        newRow.Range[1, i].NumberFormat = "$0.00";
                        break;
                    case ExcelRow.Odometer:
                    case ExcelRow.EngineHoursCounter:
                    case ExcelRow.FuelLevel:
                    case ExcelRow.DistanceTraveled:
                    case ExcelRow.EngineHoursUsed:
                    case ExcelRow.FuelUsed:
                    case ExcelRow.MPG:
                        newRow.Range[1, i].NumberFormat = "0.0";
                        break;
                    case ExcelRow.LegDuration:
                    case ExcelRow.DayDuration:
                        newRow.Range[1, i].NumberFormat = "[h]:mm:ss";
                        break;
                    default:
                        break;
                }
            }
        }

        static void InsertEventRowIntoTable(ListObject table, EventInfo e)
        {
            ListRow newRow = table.ListRows.Add();

            newRow.Range[1, ExcelRow.LocalDateTime].Value = e.LocalTime;
            newRow.Range[1, ExcelRow.TimezoneOffset].Value = e.TimeTZOffset;
            newRow.Range[1, ExcelRow.Vehicle].Value = e.VehicleName;
            if (e.GPSFixValid)
            {
                string display = $"({e.Latitude:F7}, {e.Longitude:F7})";
                string googleMapsUrl = $"https://www.google.com/maps?q={e.Latitude},{e.Longitude}";

                // Get the worksheet from the table
                Worksheet parentWorksheet = (Worksheet)table.Parent;
                parentWorksheet.Hyperlinks.Add(newRow.Range[1, ExcelRow.GPSLocation], googleMapsUrl, Type.Missing, "Open Google Maps", display);
            }

            newRow.Range[1, ExcelRow.Odometer].Value = e.Odometer;
            newRow.Range[1, ExcelRow.EngineHoursCounter].Value = e.EngineHours;
            newRow.Range[1, ExcelRow.FuelLevel].Value = e.FuelLevel;
           
            switch (e.Type)
            {
                case EventInfo.EventType.FuelPurchase:
                    newRow.Range[1, ExcelRow.Type].Value = "Gas";
                    if (e is PurchaseFuelInfo purchaseFuel)
                    {
                        newRow.Range[1, ExcelRow.Quantity].Value = purchaseFuel.Quantity;
                        newRow.Range[1, ExcelRow.Cost].Value = purchaseFuel.Cost;
                        newRow.Range[1, ExcelRow.DistanceTraveled].Value = purchaseFuel.Distance;
                        if (purchaseFuel.Quantity != 0)
                            newRow.Range[1, ExcelRow.MPG].Value = purchaseFuel.Distance / purchaseFuel.Quantity;
                    }
                    break;

                case EventInfo.EventType.PropanePurchase:
                    newRow.Range[1, ExcelRow.Type].Value = "Propane";
                    if (e is PurchasePropaneInfo purchasePropane)
                    {
                        newRow.Range[1, ExcelRow.Quantity].Value = purchasePropane.Quantity;
                        newRow.Range[1, ExcelRow.Cost].Value = purchasePropane.Cost;
                    }
                    break;

                case EventInfo.EventType.OilChange:
                    newRow.Range[1, ExcelRow.Type].Value = "Oil Change"; 
                    if (e is ChangeOilInfo changeOil)
                    {
                        newRow.Range[1, ExcelRow.DistanceTraveled].Value = changeOil.Distance;
                    }
                    break;

                default:
                    newRow.Range[1, ExcelRow.Type].Value = "Event";
                    break;
            }

            for (int i = 1; i <= ExcelRow.MaxRows; i++)
            {
                switch (i)
                {
                    case ExcelRow.TimezoneOffset:
                        newRow.Range[1, i].NumberFormat = "0";
                        break;
                    case ExcelRow.Quantity:
                        newRow.Range[1, i].NumberFormat = "0.000";
                        break;
                    case ExcelRow.Cost:
                        newRow.Range[1, i].NumberFormat = "$0.00";
                        break;
                    case ExcelRow.Odometer:
                    case ExcelRow.EngineHoursCounter:
                    case ExcelRow.FuelLevel:
                    case ExcelRow.DistanceTraveled:
                    case ExcelRow.EngineHoursUsed:
                    case ExcelRow.FuelUsed:
                    case ExcelRow.MPG:
                        newRow.Range[1, i].NumberFormat = "0.0";
                        break;
                    case ExcelRow.LegDuration:
                    case ExcelRow.DayDuration:
                        newRow.Range[1, i].NumberFormat = "[h]:mm:ss";
                        break;
                    default:
                        break;
                }
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
                    }
                    else if (_aggregateDataType == AggregateDataType.TripInfo)
                    {
                        var info = new TripInfo(_aggregateDataArray);
                        TripInfoList.Add(info);

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
                                break;

                            case EventInfo.EventType.PropanePurchase:
                                var pp = new PurchasePropaneInfo(_aggregateDataArray);
                                EventInfoList.Add(pp);
                                break;

                            case EventInfo.EventType.OilChange:
                                var oc = new ChangeOilInfo(_aggregateDataArray);
                                EventInfoList.Add(oc);
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
        public bool StartGPSFixValid { get; set; }
        public double StartLatitude { get; set; }
        public double StartLongitude { get; set; }
        public bool EndGPSFixValid { get; set; }
        public double EndLatitude { get; set; }
        public double EndLongitude { get; set; }

        public TripInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                Type = new string(reader.ReadChars(1));

                char c = reader.ReadChar();
                StartGPSFixValid = c > 0x0;

                c = reader.ReadChar();
                EndGPSFixValid = c > 0x0;

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
                StartLatitude = reader.ReadDouble();
                StartLongitude = reader.ReadDouble();
                EndLatitude = reader.ReadDouble();
                EndLongitude = reader.ReadDouble();
            }
        }

        public void Print()
        {
            Console.WriteLine($"\tStart Time: {StartTime} ({StartLocalTime}) GMT{StartTimeTZOffset}");
            Console.WriteLine($"\tEnd Time: {EndTime} ({EndLocalTime}) GMT{EndTimeTZOffset}");
            Console.WriteLine($"\tStart Odometer: {StartOdometer:F1}");
            Console.WriteLine($"\tEnd Odometer: {EndOdometer:F1}");
            Console.WriteLine($"\tStart Engine Hours: {StartEngineHours:F1}");
            Console.WriteLine($"\tEnd Engine Hours: {EndEngineHours:F1}");
            Console.WriteLine($"\tStart Fuel: {StartFuel:F1}");
            Console.WriteLine($"\tEnd Fuel: {EndFuel:F1}");
            Console.WriteLine($"\tFuel Used: {FuelUsed:F1}");
            Console.Write($"\tStart GPS {StartGPSFixValid} ");
            if (StartGPSFixValid) Console.WriteLine($"({StartLatitude:F7}, {StartLongitude:F7})"); else Console.WriteLine();
            Console.Write($"\tEnd GPS {EndGPSFixValid} ");
            if (EndGPSFixValid) Console.WriteLine($"({EndLatitude:F7}, {EndLongitude:F7})"); else Console.WriteLine();
        }
    }

    public class EventInfo
    {
        public enum EventType : byte
        {
            Unknown = 0,
            FuelPurchase = 1,
            PropanePurchase = 2,
            OilChange = 3
        }

        public uint Id { get; set; }
        public string VehicleName { get; set; }
        public EventType Type { get; set; }
        public uint Time { get; set; }
        public int TimeTZOffset { get; set; }
        public double Odometer { get; set; }
        public double EngineHours { get; set; }
        public double FuelLevel { get; set; }
        public DateTime TimeGMT { get { return new DateTime(1970, 1, 1).AddSeconds(Time); } }
        public DateTime LocalTime { get { return TimeGMT.AddHours(TimeTZOffset); } }
        public bool GPSFixValid { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public void Print()
        {
            Console.WriteLine($"\tVehicleName: {VehicleName}");
            Console.WriteLine($"\tTime: {Time} ({LocalTime}) GMT{TimeTZOffset}");
            Console.WriteLine($"\tOdometer: {Odometer:F1}");
            Console.WriteLine($"\tEngine Hours: {EngineHours:F1}");
            Console.WriteLine($"\tFuelLevel: {FuelLevel:F1}");
            Console.Write($"\tGPS {GPSFixValid} ");
            if (GPSFixValid) Console.WriteLine($"({Latitude:F7}, {Longitude:F7})"); else Console.WriteLine();
        }
    }

    public class PurchaseFuelInfo : EventInfo
    {
        public double Quantity { get; set; }
        public double Cost { get; set; }
        public double Distance { get; set; }

        public PurchaseFuelInfo(byte[] data) 
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                this.Type = (EventType)reader.ReadByte();
                
                char c = reader.ReadChar();
                this.GPSFixValid = c > 0x0;

                this.VehicleName = new string(reader.ReadChars(10));
                this.Id = reader.ReadUInt32();
                this.Time = reader.ReadUInt32();
                this.TimeTZOffset = reader.ReadInt32();
                this.Odometer = reader.ReadDouble();
                this.EngineHours = reader.ReadDouble();
                this.FuelLevel = reader.ReadDouble();
                this.Latitude = reader.ReadDouble();
                this.Longitude = reader.ReadDouble();
                this.Quantity = reader.ReadDouble();
                this.Cost = reader.ReadDouble();
                this.Distance = reader.ReadDouble();
            }
        }
        public new void Print()
        {
            Console.WriteLine($"PurchaseFuel {Id}");
            base.Print();
            Console.WriteLine($"\tQuantity: {Quantity:F3}");
            var formatedCost = Cost.ToString("C2");
            Console.WriteLine($"\tCost: {formatedCost}");
            Console.WriteLine($"\tDistance: {Distance}");
            if (Quantity != 0)
                Console.WriteLine($"\tMPG: {Distance/Quantity:F1}");
        }
    }

    public class PurchasePropaneInfo : EventInfo
    {
        public double Quantity { get; set; }
        public double Cost { get; set; }

        public PurchasePropaneInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                this.Type = (EventType)reader.ReadByte();
                
                char c = reader.ReadChar();
                this.GPSFixValid = c > 0x0;

                this.VehicleName = new string(reader.ReadChars(10));
                this.Id = reader.ReadUInt32();
                this.Time = reader.ReadUInt32();
                this.TimeTZOffset = reader.ReadInt32();
                this.Odometer = reader.ReadDouble();
                this.EngineHours = reader.ReadDouble();
                this.FuelLevel = reader.ReadDouble();
                this.Latitude = reader.ReadDouble();
                this.Longitude = reader.ReadDouble();
                this.Quantity = reader.ReadDouble();
                this.Cost = reader.ReadDouble();
            }
        }
        public new void Print()
        {
            Console.WriteLine($"PurchasePropane {Id}"); 
            base.Print();
            Console.WriteLine($"\tQuantity: {Quantity:F3}");
            var formatedCost = Cost.ToString("C2");
            Console.WriteLine($"\tCost: {formatedCost}");
        }
    }

    public class ChangeOilInfo : EventInfo
    {
        public double Distance { get; set; }
        public ChangeOilInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                this.Type = (EventType)reader.ReadByte();

                char c = reader.ReadChar();
                this.GPSFixValid = c > 0x0;

                this.VehicleName = new string(reader.ReadChars(10));
                this.Id = reader.ReadUInt32();
                this.Time = reader.ReadUInt32();
                this.TimeTZOffset = reader.ReadInt32(); 
                this.Odometer = reader.ReadDouble();
                this.EngineHours = reader.ReadDouble();
                this.FuelLevel = reader.ReadDouble();
                this.Latitude = reader.ReadDouble();
                this.Longitude = reader.ReadDouble();
                this.Distance = reader.ReadDouble();
            }
        }
        public new void Print()
        {
            Console.WriteLine($"ChangeOil {Id}");
            base.Print();
            Console.WriteLine($"\tDistance: {Distance}");
        }
    }
}
