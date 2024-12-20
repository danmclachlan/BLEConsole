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
        public string Type { get;private set; }
        public uint Id { get; private set; }
        public uint StartTime { get; private set; }
        public int StartTimeTZOffset { get; private set; }
        public DateTime StartTimeGMT { get { return new DateTime(1970, 1, 1).AddSeconds(StartTime); } }
        public DateTime StartLocalTime { get { return StartTimeGMT.AddHours(StartTimeTZOffset); } }
        public DateTime EndTimeGMT { get { return new DateTime(1970, 1, 1).AddSeconds(EndTime); } }
        public DateTime EndLocalTime { get { return EndTimeGMT.AddHours(EndTimeTZOffset); } }
        public uint EndTime { get; private set; }
        public int EndTimeTZOffset { get; private set; }
        public double StartOdometer { get; private set; }
        public double EndOdometer { get; private set; }
        public double StartEngineHours { get; private set; }
        public double EndEngineHours { get; private set; }
        public double StartGenHrs { get; private set; }
        public double EndGenHrs { get; private set; }
        public double StartFuel { get; private set; }
        public double EndFuel { get; private set; }
        public double FuelUsed { get; private set; }
        public bool StartGPSFixValid { get; private set; }
        public double StartLatitude { get; private set; }
        public double StartLongitude { get; private set; }
        public bool EndGPSFixValid { get; private set; }
        public double EndLatitude { get; private set; }
        public double EndLongitude { get; private set; }
        public uint TravelDurationSecs { get; private set; }
        public bool IsTowing { get; private set; }
        public double TowingDistance { get; private set; }
        
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
