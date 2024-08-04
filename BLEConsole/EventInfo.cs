using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLEConsole
{
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
                Console.WriteLine($"\tMPG: {Distance / Quantity:F1}");
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

