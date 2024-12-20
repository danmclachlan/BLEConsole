using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace BLEConsole
{
    public class Event2Info
    {
        public enum EventType : byte
        {
            StartDay = 1,
            EndDay = 2,
            StartLeg = 3,
            EndLeg = 4,
            Gas = 5,
            Propane = 6,
            OilChange = 7,
            GeneratorOilChange = 8
        }

        public uint InterfaceVersion { get; internal set; }
        public uint Id { get; internal set; }
        public EventType Type { get; internal set; }
        public uint TimeSeconds { get; internal set; }
        public int TZOffset { get; internal set; }
        public uint TimeGMTSeconds { get; internal set; }
        public DateTime DateTimeLocal { get { return new DateTime(1970, 1, 1).AddSeconds(TimeSeconds); ; } }
        public DateTime DateTimeGMT { get { return new DateTime(1970, 1, 1).AddSeconds(TimeGMTSeconds); } }
        public string VehicleName { get; internal set; }
        public double Odometer { get; internal set; }
        public double EngineHours { get; internal set; }
        public double FuelLevel { get; internal set; }
        public string Description { get; internal set; }
        public bool IsTowing { get; internal set; }
        public string TowVehicle { get; internal set; }
        public bool GPSFixValid { get; internal set; }
        public double Latitude { get; internal set; }
        public double Longitude { get; internal set; }
        public bool HasGenerator { get; internal set; }
        public double GenHrsCounter { get; internal set; }

        public void Print()
        {
            Console.WriteLine($"\tTime: {DateTimeLocal} GMT{TZOffset} ({DateTimeGMT}) ");
            Console.WriteLine($"\tVehicleName: {VehicleName}");
            Console.WriteLine($"\tOdometer: {Odometer:F1}");
            Console.WriteLine($"\tEngine Hours: {EngineHours:F1}");
            Console.WriteLine($"\tFuelLevel: {FuelLevel:F1}");
            if (Description != null) Console.WriteLine($"\tDescription: {Description}");
            Console.Write($"\tTow: {IsTowing} ");
            if (IsTowing) Console.WriteLine($"\tTowVehicle: {TowVehicle}"); else Console.WriteLine();
            Console.Write($"\tGPS {GPSFixValid} ");
            if (GPSFixValid) Console.WriteLine($"({Latitude:F7}, {Longitude:F7})"); else Console.WriteLine();
            Console.Write($"\tGenerator: {HasGenerator} ");
            if (HasGenerator) Console.WriteLine($"{GenHrsCounter}"); else Console.WriteLine();
        }

        protected void InitializeCommonData(BinaryReader reader)
        {
            this.InterfaceVersion = reader.ReadUInt32();
            this.Id = reader.ReadUInt32();
            this.Type = (EventType)reader.ReadByte();
            this.TimeSeconds = reader.ReadUInt32();
            this.TZOffset = reader.ReadInt32();
            this.TimeGMTSeconds = reader.ReadUInt32();
            this.VehicleName = new string(reader.ReadChars(10));
            this.Odometer = reader.ReadDouble();
            this.EngineHours = reader.ReadDouble();
            this.FuelLevel = reader.ReadDouble();

            uint descLength = reader.ReadUInt32();
            if (descLength > 0)
            {
                this.Description = new string(reader.ReadChars((int)descLength));
            }

            char c = reader.ReadChar();
            this.IsTowing = c > 0x0;
            if (IsTowing)
            {
                this.TowVehicle = new string(reader.ReadChars(10));
            }

            c = reader.ReadChar();
            this.GPSFixValid = c > 0x0;
            if (GPSFixValid)
            {
                this.Latitude = reader.ReadDouble();
                this.Longitude = reader.ReadDouble();
            }

            c = reader.ReadChar();
            this.HasGenerator = c > 0x0;
            if (HasGenerator)
                this.GenHrsCounter = reader.ReadDouble();
        }
    }

    public class StartDayInfo : Event2Info
    {
        public StartDayInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                InitializeCommonData(reader);
            }
        }

        public new void Print()
        {
            Console.WriteLine($"Start Day {Id}");
            base.Print();
        }
    }

    public class StartLegInfo : Event2Info
    {
        public StartLegInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                InitializeCommonData(reader);
            }
        }

        public new void Print()
        {
            Console.WriteLine($"Start Leg {Id}");
            base.Print();
        }
    }

    public class EndLegInfo : Event2Info
    {
        public double EngineHoursUsed { get; internal set; }
        public double FuelUsed { get; internal set; }
        public double TowingDistance { get; internal set; }
        public double Distance { get; internal set; }
        public uint DurationSeconds { get; internal set; }
        public TimeSpan Duration { get { return TimeSpan.FromSeconds(DurationSeconds); } }
        public double AvgMPH { get; internal set; }

        public EndLegInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                InitializeCommonData(reader);

                // EndLeg & EndDay Specific Data
                this.EngineHoursUsed = reader.ReadDouble();
                this.FuelUsed = reader.ReadDouble();
                this.TowingDistance = reader.ReadDouble();
                this.Distance = reader.ReadDouble();
                this.DurationSeconds = reader.ReadUInt32();

                // EndLeg Specific Data
                this.AvgMPH = reader.ReadDouble();
            }
        }

        public new void Print()
        {
            Console.WriteLine($"End Leg {Id}");
            base.Print();
            Console.WriteLine();
            Console.WriteLine($"\tEngine Hours Used: {EngineHoursUsed}");
            Console.WriteLine($"\tFuel Used: {FuelUsed}");
            Console.WriteLine($"\tDistance: {Distance}");
            Console.WriteLine($"\tTowing Distance: {TowingDistance}");
            Console.WriteLine($"\tDuration: {Duration}");
            Console.WriteLine($"\tAvg MPH: {AvgMPH}");
        }
    }

    public class EndDayInfo : Event2Info
    {
        public double EngineHoursUsed { get; internal set; }
        public double FuelUsed { get; internal set; }
        public double TowingDistance { get; internal set; }
        public double Distance { get; internal set; }
        public uint DurationSeconds { get; internal set; }
        public TimeSpan Duration { get { return TimeSpan.FromSeconds(DurationSeconds); } }
        public uint TravelDurationSeconds { get; internal set; }
        public TimeSpan TravelDuration { get { return TimeSpan.FromSeconds(TravelDurationSeconds); } }

        public EndDayInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                InitializeCommonData(reader);

                // EndLeg & EndDay Specific Data
                this.EngineHoursUsed = reader.ReadDouble();
                this.FuelUsed = reader.ReadDouble();
                this.TowingDistance = reader.ReadDouble();
                this.Distance = reader.ReadDouble();
                this.DurationSeconds = reader.ReadUInt32();

                // EndDay Specific Data
                this.TravelDurationSeconds = reader.ReadUInt32();
            }
        }

        public new void Print()
        {
            Console.WriteLine($"End Day {Id}");
            base.Print();
            Console.WriteLine();
            Console.WriteLine($"\tEngine Hours Used: {EngineHoursUsed}");
            Console.WriteLine($"\tFuel Used: {FuelUsed}");
            Console.WriteLine($"\tDistance: {Distance}");
            Console.WriteLine($"\tTowing Distance: {TowingDistance}");
            Console.WriteLine($"\tDuration: {Duration}");
            Console.WriteLine($"\tTravel Duration: {TravelDuration}");
        }
    }

    public class GasInfo : Event2Info
    {
        public double Quantity { get; internal set; }
        public double Cost { get; internal set; }
        public double Distance { get; internal set; }

        public GasInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                InitializeCommonData(reader);

                // Gas Info Specific Data
                this.Quantity = reader.ReadDouble();
                this.Cost = reader.ReadDouble();
                this.Distance = reader.ReadDouble();
            }
        }

        public new void Print()
        {
            Console.WriteLine($"Gas {Id}");
            base.Print();
            Console.WriteLine($"\tQuantity: {Quantity:F3}");
            var formatedCost = Cost.ToString("C2");
            Console.WriteLine($"\tCost: {formatedCost}");
            Console.WriteLine($"\tDistance: {Distance}");
            if (Quantity != 0)
                Console.WriteLine($"\tMPG: {Distance / Quantity:F1}");
        }
    }

    public class PropaneInfo : Event2Info
    {
        public double Quantity { get; internal set; }
        public double Cost { get; internal set; }

        public PropaneInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                InitializeCommonData(reader);

                // Propane Info Specific Data
                this.Quantity = reader.ReadDouble();
                this.Cost = reader.ReadDouble();
            }
        }

        public new void Print()
        {
            Console.WriteLine($"Propane {Id}");
            base.Print();
            Console.WriteLine($"\tQuantity: {Quantity:F3}");
            var formatedCost = Cost.ToString("C2");
            Console.WriteLine($"\tCost: {formatedCost}");
        }
    }

    public class OilChangeInfo : Event2Info
    {
        public double Distance { get; internal set; }

        public OilChangeInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                InitializeCommonData(reader);

                // Oil Change Info Specific Data
                this.Distance = reader.ReadDouble();
            }
        }

        public new void Print()
        {
            Console.WriteLine($"Change Oil {Id}");
            base.Print();
            Console.WriteLine($"\tDistance: {Distance}");
        }
    }

    public class GeneratorOilChangeInfo : Event2Info
    {
        public GeneratorOilChangeInfo(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data))
            using (BinaryReader reader = new BinaryReader(stream))
            {
                InitializeCommonData(reader);
            }
        }
        public new void Print()
        {
            Console.WriteLine($"Generator Change Oil {Id}");
            base.Print();
            Console.WriteLine($"\tHours: {GenHrsCounter}");
        }
    }
}