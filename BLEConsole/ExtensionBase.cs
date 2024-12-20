using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth.GenericAttributeProfile;

namespace BLEConsole
{
    internal abstract class ExtensionBase
    {
        public abstract Task<(bool, int)> ExecuteExtensionAsync(string cmd, string parameters);
        public abstract bool Characteristic_ValueChanged(GattCharacteristic sender, GattValueChangedEventArgs args);
        public abstract void Help();
    }
}
