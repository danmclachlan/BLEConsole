using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLEConsole
{
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
        public bool IsAbleToTow {  get; set; }
        public string TowVehicle {  get; set; }
        public bool HasGenerator { get; set; }
        public double GeneratorHours { get; set; }

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
                this.IsAbleToTow = reader.ReadByte() == 0x1;
                if (this.IsAbleToTow)
                    this.TowVehicle = new string(reader.ReadChars(10));
                this.HasGenerator = reader.ReadByte() == 0x1;
                if (this.HasGenerator)
                    this.GeneratorHours = reader.ReadDouble();
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
            if ( IsAbleToTow ) 
                Console.Write($"\tTow Vehicle: {TowVehicle}\n");
            if (HasGenerator)
                Console.Write($"\tGenerator Hours: {GeneratorHours}\n");
        }
    }
}
