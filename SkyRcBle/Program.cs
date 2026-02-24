// ==============================================================================
// Copyright (c) 2026 V.Loos (umodified MIT license)
// Permission is hereby granted, free of charge, to any person obtaining a copy of
// this software and associated documentation files (the “Software”), to deal in the
// Software without restriction, including without limitation the rights to use, copy,
// modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
// and to permit persons to whom the Software is furnished to do so, subject to the
// following conditions:
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
// PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
// HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
// SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
// ==============================================================================

using System;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.Advertisement;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Enumeration;


namespace BtExperiment
{
    class Program
    {
        private static List<BleCharger> _devs = new List<BleCharger>();
        private static readonly System.Collections.Generic.Dictionary<ulong, string> DeviceList = new();

        static async Task Main(string[] args)
        {
            Console.WriteLine("Press ESC during sampling to stop sampling and disconnect.");

            // search for connected devices
            string selector = BluetoothLEDevice.GetDeviceSelectorFromConnectionStatus(BluetoothConnectionStatus.Connected); // BLE only
            var devices = await DeviceInformation.FindAllAsync(selector);
            string result = $"At start there were {devices.Count} connected device(s)\n";
            foreach (var dev in devices)
            {
                BluetoothLEDevice? ble = null;
                try
                {
                    ble = await BluetoothLEDevice.FromIdAsync(dev.Id);
                }
                catch { }
                result += $" Name: {dev.Name}, ID: {dev.Id}, Kind: {dev.Kind}, Conn: {ble?.ConnectionStatus}, Paired: {dev.Pairing?.CanPair}\n";

                if (dev.Name == "Charger")  // Windows has that name from GAP service (so available only if connected).
                {
                    string macStr = dev.Id.Split('-')[1];
                    ulong mac = Convert.ToUInt64(macStr.Replace(":", ""), 16);
                    BleCharger bleDev = new BleCharger(mac);
                    await bleDev.ConnectRequest(); // naming does not fit here (already connected)
                    _devs.Add(bleDev);
                }
            }
            Console.WriteLine(result);

            // Scan for BLE deivces and connects to all Chargers.
            int scanTime_s = 20;
            if (_devs.Count == 0) // TODO: better decision if scan is needed.
            {
                Console.WriteLine($"Starting BLE scan for {scanTime_s} seconds...");
                BluetoothLEAdvertisementWatcher _watcher;
                _watcher = new BluetoothLEAdvertisementWatcher
                {
                    ScanningMode = BluetoothLEScanningMode.Passive // otherwise heavy interference with Bluetooth Classic devices
                };
                _watcher.Received += OnAdvertisementReceived;
                _watcher.Stopped += OnAdvertisementWatcherStopped;
                _watcher.Start();
                await Task.Delay(scanTime_s * 1000);
                _watcher.Stop();


                foreach (var dev in DeviceList)
                {
                    string name = dev.Value;
                    if (dev.Value.Contains("SkyRC"))
                    {
                        BleCharger bleDev = new BleCharger(dev.Key);
                        await bleDev.ConnectRequest();
                        _devs.Add(bleDev);
                    }
                }
            }

            // Or connect to known charger(s). Only if the device is in Windows cache (was seen during a scan).
            //ulong mac1 = 0x123456789ABC;
            //BleCharger bleDev1 = new BleCharger(mac1);
            //await bleDev1.ConnectAndSetupAsync();
            //_devs.Add(bleDev1);

            if (_devs.Count == 0)
            {
                Console.WriteLine("Scan finished. No devices found. Exit.");
                return;
            }

            while (true)
            {
                await Task.Delay(1000);
                if (Console.KeyAvailable)
                {
                    ConsoleKeyInfo key = Console.ReadKey(intercept: true);
                    if (key.Key == ConsoleKey.Escape)
                        break;
                    else if (key.Key == ConsoleKey.R)
                    {
                        foreach (BleCharger dev in _devs)
                        {
                            dev.RestartCsv();
                        }
                    }
                }
            }

            Console.WriteLine("Disconnecting devices...");
            foreach (BleCharger dev in _devs)
            {
                await dev.Disconnect();
            }
            Console.WriteLine("Done.");
            await Task.Delay(1000);
        }

        // Analyze only the first received advertisement paket from each device. Add an entry in DeviceList (MAC and a free defined name). 
        private static void OnAdvertisementReceived(BluetoothLEAdvertisementWatcher watcher, BluetoothLEAdvertisementReceivedEventArgs args)
        {
            string deviceName = args.Advertisement.LocalName;

            if (!DeviceList.ContainsKey(args.BluetoothAddress)) 
            {
                // Try to identify SkyRC MC3000 chargers (just a guess, make break on update)
                foreach (var mfg in args.Advertisement.ManufacturerData)
                {
                    byte[] mfgData = mfg.Data.ToArray();
                    string hex = BitConverter.ToString(mfgData).Replace("-", " ");
                    if (mfg.CompanyId == (args.BluetoothAddress & 0xFFFF))  // Not a company ID. Just 16bit from the MAC address (observed with SkyRC MC3000, firmware version 1.21 and 1.25 of)
                    {
                        deviceName = "SkyRC?";
                    }
                    if (mfgData.Length == 20 && mfgData[4] == 0x31 && mfgData[5] == 0x30 && mfgData[6] == 0x30 && mfgData[7] == 0x30 && mfgData[8] == 0x38 && mfgData[9] == 0x33) // 31 30 30 30 38 33 (same for firmware version 1.21 and 1.25 of SkyRC MC3000)
                    {
                        deviceName = "SkyRC?";
                    }
                }

                // List all advertising devices
                Console.Write($"{args.BluetoothAddress:X12}; {args.RawSignalStrengthInDBm} dBm; ");
                foreach (var section in args.Advertisement.DataSections)
                {
                    byte[] data = section.Data.ToArray();
                    string hex = BitConverter.ToString(data).Replace("-", " ");
                    Console.Write($"({data.Length}) {section.DataType:X2} {hex}; ");
                }
                if (deviceName.Length > 0)
                    Console.WriteLine($"\"{deviceName}\""); // Many devices do not have a name in advertisement packets.
                else
                    Console.WriteLine("");

                DeviceList[args.BluetoothAddress] = deviceName;
            }
        }

        private static void OnAdvertisementWatcherStopped(BluetoothLEAdvertisementWatcher watcher, BluetoothLEAdvertisementWatcherStoppedEventArgs args)
        {
            Console.WriteLine("Scan stopped.");
        }
    }
}
