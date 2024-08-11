using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLEConsole
{
    public class TripInfo
    {
        public string Type { get; set; }
        public uint Id { get; set; }
        public uint StartTime { get; set; }
        public int StartTimeTZOffset { get; set; }
        public DateTime StartTimeGMT { get { return new DateTime(1970, 1, 1).AddSeconds(StartTime); } }
        public DateTime StartLocalTime { get { return StartTimeGMT.AddHours(StartTimeTZOffset); } }
        public DateTime EndTimeGMT { get { return new DateTime(1970, 1, 1).AddSeconds(EndTime); } }
        public DateTime EndLocalTime { get { return EndTimeGMT.AddHours(EndTimeTZOffset); } }
        public uint EndTime { get; set; }
        public int EndTimeTZOffset { get; set; }
        public double StartOdometer { get; set; }
        public double EndOdometer { get; set; }
        public double StartEngineHours { get; set; }
        public double EndEngineHours { get; set; }
        public double StartGenHrs { get; set; }
        public double EndGenHrs { get; set; }
        public double StartFuel { get; set; }
        public double EndFuel { get; set; }
        public double FuelUsed { get; set; }
        public bool StartGPSFixValid { get; set; }
        public double StartLatitude { get; set; }
        public double StartLongitude { get; set; }
        public bool EndGPSFixValid { get; set; }
        public double EndLatitude { get; set; }
        public double EndLongitude { get; set; }
        public uint TravelDurationSecs { get; set; }
        public bool IsTowing { get; set; }
        public double TowingDistance { get; set; }
        
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
                StartGenHrs = reader.ReadDouble();
                EndGenHrs = reader.ReadDouble();
                StartFuel = reader.ReadDouble();
                EndFuel = reader.ReadDouble();
                FuelUsed = reader.ReadDouble();
                StartLatitude = reader.ReadDouble();
                StartLongitude = reader.ReadDouble();
                EndLatitude = reader.ReadDouble();
                EndLongitude = reader.ReadDouble();
                TravelDurationSecs = reader.ReadUInt32();
                IsTowing = reader.ReadByte() == 0x1;
                TowingDistance = reader.ReadDouble();
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
            Console.WriteLine($"\tStart Generator Hours: {StartGenHrs}");
            Console.WriteLine($"\tEnd Generator Hours: {EndGenHrs}");
            Console.WriteLine($"\tStart Fuel: {StartFuel:F1}");
            Console.WriteLine($"\tEnd Fuel: {EndFuel:F1}");
            Console.WriteLine($"\tFuel Used: {FuelUsed:F1}");
            Console.Write($"\tStart GPS {StartGPSFixValid} ");
            if (StartGPSFixValid) Console.WriteLine($"({StartLatitude:F7}, {StartLongitude:F7})"); else Console.WriteLine();
            Console.Write($"\tEnd GPS {EndGPSFixValid} ");
            if (EndGPSFixValid) Console.WriteLine($"({EndLatitude:F7}, {EndLongitude:F7})"); else Console.WriteLine();
            double durationHrs = TravelDurationSecs / 3600.0;
            Console.WriteLine($"\tTravel Duration: {TravelDurationSecs} secs ({durationHrs:F3} hrs)");
            Console.WriteLine($"\tIs Towing: {IsTowing}");
            Console.WriteLine($"\tTowing Distance: {TowingDistance}");
        }
    }
}
