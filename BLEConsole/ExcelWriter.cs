using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLEConsole
{
    internal class ExcelWriter : IDisposable
    {
        internal string FilePath { get; set; }
        internal Application ExcelApp { get; set; }
        internal Workbook CurrentWB { get; set; }
        internal Worksheet ConfigSheet { get; set; }
        internal ListObject TripDetailTable { get; set; }
        internal int TripId { get; set; }

        internal enum ExcelInsertType { DayStart, LegStart, LegEnd, DayEnd };

        public ExcelWriter(string filePath)
        {
            FilePath = filePath;

            ExcelApp = new Application();
            CurrentWB = ExcelApp.Workbooks.Open(FilePath);

            if (CurrentWB != null)
            {
                ConfigSheet = CurrentWB.Sheets["Config"] as Worksheet;

                TripDetailTable = GetTableByName("Travel_Log_All");

                // Get the value of the current trip id.
                Name namedRange = CurrentWB.Names.Item("CurrentTrip");
                Range range = namedRange?.RefersToRange;
                TripId = (int)range?.Value2;
            }
            //ExcelApp.Visible = true;
        }

        public void Dispose()
        {
            CurrentWB?.Close(true);
            ExcelApp.Quit();
            if (CurrentWB != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(CurrentWB);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
        }

        public void AppendToTripDetailTable(VehicleInfo vehicle, List<TripInfo> tripInfo, List<EventInfo> eventInfo)
        {
            if (TripDetailTable == null)
            {
                Console.WriteLine($"Error: TripDetailTable not defined");
                return;
            }

            bool locked = IsListObjectProtected(TripDetailTable);

            if (locked)
            {
                SetSheetProtection((Worksheet)TripDetailTable.Parent, false);
            }

            // Turn off calculations while adding the data to sped up the process.
            ExcelApp.Calculation = Excel.XlCalculation.xlCalculationManual;

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
            for (int i = 0; i < tripInfo.Count; i++)
            {
                // Find any events that occurred before the start of this Day/Leg
                if (j < eventInfo.Count && eventInfo[j].Time < tripInfo[i].StartTime)
                    InsertEventRowIntoTable(eventInfo[j++]);

                if (i == 0)
                    InsertTripRowIntoTable(vehicle, ExcelInsertType.DayStart, tripInfo[i]);
                else
                {
                    InsertTripRowIntoTable(vehicle, ExcelInsertType.LegStart, tripInfo[i]);
                    InsertTripRowIntoTable(vehicle, ExcelInsertType.LegEnd, tripInfo[i]);
                }
            }
            // Find any remaining event that occurred after the end of the last leg
            while (j < eventInfo.Count)
                InsertEventRowIntoTable(eventInfo[j++]);

            InsertTripRowIntoTable(vehicle, ExcelInsertType.DayEnd, tripInfo[0]);

            // Reenable automatic calculations once the data has been added
            ExcelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

            if (locked)
            {
                SetSheetProtection((Worksheet)TripDetailTable.Parent, true);
            }
            CurrentWB?.Save();
        }

        internal void InsertTripRowIntoTable(VehicleInfo vehicle,
            ExcelInsertType insertType, TripInfo trip)
        {
            ListRow newRow = TripDetailTable.ListRows.Add();
            SetFormatOnColumns(newRow);

            newRow.Range[1, ExcelRow.TripId] = TripId;
            if (vehicle.Name != null) newRow.Range[1, ExcelRow.Vehicle].Value = vehicle.Name;
            if (vehicle.IsAbleToTow && trip.IsTowing) newRow.Range[1, ExcelRow.TowVehicle].Value = vehicle.TowVehicle;

            switch (insertType)
            {
                case ExcelInsertType.DayStart:
                    newRow.Range[1, ExcelRow.Type].Value = "S";
                    break;

                case ExcelInsertType.DayEnd:
                    newRow.Range[1, ExcelRow.Type].Value = "E";
                    break;

                case ExcelInsertType.LegStart:
                    newRow.Range[1, ExcelRow.Type].Value = "D";
                    break;

                case ExcelInsertType.LegEnd:
                    newRow.Range[1, ExcelRow.Type].Value = "A";
                    break;
            }

            switch (insertType)
            {
                case ExcelInsertType.DayStart:
                case ExcelInsertType.LegStart:
                    newRow.Range[1, ExcelRow.LocalDate].Value = trip.StartLocalTime.ToShortDateString();
                    newRow.Range[1, ExcelRow.LocalTime].Value = trip.StartLocalTime.ToLongTimeString();
                    newRow.Range[1, ExcelRow.TimezoneOffset] = trip.StartTimeTZOffset;
                    newRow.Range[1, ExcelRow.TimeGMT].Value = trip.StartTimeGMT;

                    if (trip.StartGPSFixValid)
                        InsertGPSDataHyperLink(newRow, trip.StartLatitude, trip.StartLongitude);
                    
                    newRow.Range[1, ExcelRow.Odometer].Value = trip.StartOdometer;
                    newRow.Range[1, ExcelRow.EngineHoursCounter].Value = trip.StartEngineHours;
                    newRow.Range[1, ExcelRow.FuelLevel].Value = trip.StartFuel;

                    if (vehicle.HasGenerator)
                        newRow.Range[1, ExcelRow.GenHrs].Value = trip.StartGenHrs;
                    break;

                case ExcelInsertType.DayEnd:
                case ExcelInsertType.LegEnd:
                    double distTraveled = trip.EndOdometer - trip.StartOdometer;
                    TimeSpan legDuration = trip.EndTimeGMT - trip.StartTimeGMT;

                    newRow.Range[1, ExcelRow.LocalDate].Value = trip.EndLocalTime.ToShortDateString();
                    newRow.Range[1, ExcelRow.LocalTime].Value = trip.EndLocalTime.ToLongTimeString();
                    newRow.Range[1, ExcelRow.TimezoneOffset] = trip.EndTimeTZOffset;
                    newRow.Range[1, ExcelRow.TimeGMT].Value = trip.EndTimeGMT;

                    if (trip.EndGPSFixValid)
                        InsertGPSDataHyperLink(newRow, trip.EndLatitude, trip.EndLongitude);
                    
                    newRow.Range[1, ExcelRow.Odometer].Value = trip.EndOdometer;
                    newRow.Range[1, ExcelRow.EngineHoursCounter].Value = trip.EndEngineHours;
                    newRow.Range[1, ExcelRow.FuelLevel].Value = trip.EndFuel;

                    if (vehicle.HasGenerator)
                        newRow.Range[1, ExcelRow.GenHrs].Value = trip.EndGenHrs;

                    newRow.Range[1, ExcelRow.EngineHrsUsed].Value = trip.EndEngineHours - trip.StartEngineHours;
                    newRow.Range[1, ExcelRow.FuelUsed].Value = trip.FuelUsed;
                    if (trip.FuelUsed > 0)
                        newRow.Range[1, ExcelRow.MPG].Value = distTraveled / trip.FuelUsed;

                    if (vehicle.IsAbleToTow && vehicle.TowVehicle != null)
                        newRow.Range[1, ExcelRow.TowedMilesPerDay].Value = trip.TowingDistance;

                    if (insertType == ExcelInsertType.LegEnd)
                    {
                        // End of Leg
                        // legDuration ==> TimePerLeg
                        // distTraveled ==> DistPerLeg
                        newRow.Range[1, ExcelRow.TimePerLeg].Value = legDuration.TotalSeconds / 86400; // Excel stores time as a fraction of a day
                        newRow.Range[1, ExcelRow.DistPerLeg].Value = distTraveled;
                        if (legDuration.TotalSeconds > 0)
                            newRow.Range[1, ExcelRow.AvgMPH].Value = distTraveled / legDuration.TotalHours;
                    }
                    else
                    {
                        // End of the Day 
                        // legDuration ==> ElapsedTimePerDay
                        // distTraveled ==> TotalDistPerDay
                        // LegDurationSum ==> TimeTravelingPerDay
                        newRow.Range[1, ExcelRow.ElapsedTimePerDay].Value = legDuration.TotalSeconds / 86400; // Excel stores time as a fraction of a day
                        newRow.Range[1, ExcelRow.TotalDistPerDay].Value = distTraveled;
                        newRow.Range[1, ExcelRow.TimeTravelingPerDay].Value = (double)(trip.TravelDurationSecs) / 86400; // Excel stores time as a fraction of a day
                    }
                    break;
            }
        }

        internal void InsertEventRowIntoTable(EventInfo e)
        {
            ListRow newRow = TripDetailTable.ListRows.Add();
            SetFormatOnColumns(newRow);

            newRow.Range[1, ExcelRow.TripId] = TripId;
            if (e.VehicleName != null) newRow.Range[1, ExcelRow.Vehicle].Value = e.VehicleName;

            newRow.Range[1, ExcelRow.LocalDate].Value = e.LocalTime.ToShortDateString();
            newRow.Range[1, ExcelRow.LocalTime].Value = e.LocalTime.ToLongTimeString();
            newRow.Range[1, ExcelRow.TimezoneOffset] = e.TimeTZOffset;
            newRow.Range[1, ExcelRow.TimeGMT].Value = e.TimeGMT;

            
            if (e.GPSFixValid)
                InsertGPSDataHyperLink(newRow, e.Latitude, e.Longitude);

            newRow.Range[1, ExcelRow.Odometer].Value = e.Odometer;
            newRow.Range[1, ExcelRow.EngineHoursCounter].Value = e.EngineHours;
            newRow.Range[1, ExcelRow.GenHrs].Value = e.GenHrs;
            newRow.Range[1, ExcelRow.FuelLevel].Value = e.FuelLevel;

            switch (e.Type)
            {
                case EventInfo.EventType.FuelPurchase:
                    if (e.VehicleName == "F09")
                    {
                        newRow.Range[1, ExcelRow.Cash].Value = "C";
                    }
                    else
                    {
                        newRow.Range[1, ExcelRow.Cash].Value = "G";
                    }
                    if (e is PurchaseFuelInfo purchaseFuel)
                    {
                        newRow.Range[1, ExcelRow.Quantity].Value = purchaseFuel.Quantity;
                        newRow.Range[1, ExcelRow.Cost].Value = purchaseFuel.Cost;
                        newRow.Range[1, ExcelRow.DistPerLeg].Value = purchaseFuel.Distance;
                        if (purchaseFuel.Quantity != 0)
                            newRow.Range[1, ExcelRow.MPG].Value = purchaseFuel.Distance / purchaseFuel.Quantity;
                    }
                    break;

                case EventInfo.EventType.PropanePurchase:
                    newRow.Range[1, ExcelRow.Cash].Value = "P";
                    if (e is PurchasePropaneInfo purchasePropane)
                    {
                        newRow.Range[1, ExcelRow.Quantity].Value = purchasePropane.Quantity;
                        newRow.Range[1, ExcelRow.Cost].Value = purchasePropane.Cost;
                    }
                    break;

                case EventInfo.EventType.OilChange:
                    newRow.Range[1, ExcelRow.Cash].Value = "R";
                    newRow.Range[1, ExcelRow.MaintenanceType] = "Oil";
                    if (e is ChangeOilInfo changeOil)
                    {
                        newRow.Range[1, ExcelRow.DistPerLeg].Value = changeOil.Distance;
                    }
                    break;
            }
        }

        internal void InsertGPSDataHyperLink(ListRow row, double lat, double lon)
        {
            string display = $"({lat:F7}, {lon:F7})";
            string googleMapsUrl = $"https://www.google.com/maps?q={lat},{lon}";

            // Get the worksheet from the table
            Worksheet parentWorksheet = (Worksheet)TripDetailTable.Parent;
            parentWorksheet.Hyperlinks.Add(row.Range[1, ExcelRow.GPSLocation], googleMapsUrl, Type.Missing, "Open Google Maps", display);
        }

        public static class ExcelRow
        {
            public const int TripId = 1;
            public const int LocalDate = 2;
            public const int LocalTime = 3;
            public const int TimezoneOffset = 4;
            public const int Vehicle = 5;               // string
            public const int TowVehicle = 6;            // string
            public const int Quantity = 7;
            public const int Cost = 8;
            public const int Cash = 9;                  // R - RV Service (See MaintenanceType), G - RV Gas, P - RV Propane, O - Other, L - Camping, E - Entertainment, C - Car, F - Food
            public const int Type = 10;                 // D - Depart, A - Arrive, S - Start, E - End
            public const int GPSLocation = 11;          // hyperlinked string
            public const int Description = 12;          // string
            public const int MaintenanceType = 13;      // Maint, Oil, Gen-Oil, Repair, Feature
            public const int Odometer = 14;
            public const int Miles = 15;                // Trip counter
            public const int EngineHoursCounter = 16;
            public const int FuelLevel = 17;
            public const int GenHrs = 18;
            public const int MilesSinceLastGas = 19;
            public const int MPG = 20;
            public const int TimeGMT = 21;              // Date and Time
            public const int DistPerLeg = 22;
            public const int TimePerLeg = 23;
            public const int EngineHrsUsed = 24;
            public const int FuelUsed = 25;
            public const int AvgMPH = 26;
            public const int TotalDistPerDay = 27;
            public const int TimeTravelingPerDay = 28;
            public const int ElapsedTimePerDay = 29;
            public const int TowedMilesPerDay = 30;

            public const int MaxRows = 30;
        }

        internal ListObject GetTableByName(string name)
        {
            if (CurrentWB != null)
            {
                foreach (Worksheet sheet in CurrentWB.Sheets)
                {
                    foreach (ListObject tbl in sheet.ListObjects)
                    {
                        if (tbl.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                            return tbl;
                    }
                }
            }
            return null;
        }

        static internal bool IsListObjectProtected(ListObject tbl)
        {
            Worksheet parentWorksheet = (Worksheet)tbl.Parent;
            return parentWorksheet.ProtectContents;
        }

        internal void SetSheetProtection(Worksheet ws, bool protect)
        {
            if (ConfigSheet != null)
            {
                string password = ConfigSheet.Cells[3, 8].Value2.ToString();

                if (password != null)
                {
                    if (protect)
                    {
                        ws.Protect(
                            Password: password,
                            DrawingObjects: true,
                            Contents: true,
                            Scenarios: true,
                            AllowFiltering: false,
                            AllowFormattingRows: true,
                            AllowFormattingColumns: true,
                            AllowInsertingRows: true
                        );
                    }
                    else
                    {
                        ws.Unprotect(password);
                    }
                }
            }
        }
        internal void SetFormatOnColumns(ListRow row)
        {
            // Clear out the analysis cells that have formulas since we are now filling in data that has already been 
            // computed on the TripTracking device.
            {
                Worksheet ws = (Worksheet)TripDetailTable.Parent;
                Range analysisRange = ws.Range[ws.Cells[row.Index+1, ExcelRow.MPG], ws.Cells[row.Index+1, ExcelRow.TowedMilesPerDay]];
                analysisRange.ClearContents();
            }

            // Format the cells in the added row to ensure they are formatted as expected.
            for (int i = 1; i <= ExcelRow.MaxRows; i++)
            {
                if (i == ExcelRow.Description)
                {
                    row.Range[1, i].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                }
                else
                {
                    row.Range[1, i].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                row.Range[1, i].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                switch (i)
                {
                    case ExcelRow.TripId:
                    case ExcelRow.TimezoneOffset:
                        row.Range[1, i].NumberFormat = "0";
                        break;
                    case ExcelRow.LocalDate:
                        row.Range[1, i].NumberFormat = "mm/dd/yy";
                        break;
                    case ExcelRow.LocalTime:
                        row.Range[1, i].NumberFormat = "hh:mm:ss";
                        break;
                    case ExcelRow.TimeGMT:
                        row.Range[1, i].NumberFormat = "mm/dd/yy hh:mm:ss";
                        break;
                    case ExcelRow.Quantity:
                        row.Range[1, i].NumberFormat = "0.000";
                        break;
                    case ExcelRow.Cost:
                        row.Range[1, i].NumberFormat = "$0.00";
                        break;
                    case ExcelRow.Odometer:
                    case ExcelRow.EngineHoursCounter:
                    case ExcelRow.FuelLevel:
                    case ExcelRow.GenHrs:
                    case ExcelRow.MilesSinceLastGas:
                    case ExcelRow.DistPerLeg:
                    case ExcelRow.EngineHrsUsed:
                    case ExcelRow.FuelUsed:
                    case ExcelRow.MPG:
                    case ExcelRow.AvgMPH:
                    case ExcelRow.TotalDistPerDay:
                    case ExcelRow.TowedMilesPerDay:
                        row.Range[1, i].NumberFormat = "0.0";
                        break;
                    case ExcelRow.TimePerLeg:
                    case ExcelRow.TimeTravelingPerDay:
                    case ExcelRow.ElapsedTimePerDay:
                        row.Range[1, i].NumberFormat = "[h]:mm:ss";
                        break;
                    default:
                        break;
                }
            }

        }

    }
}
