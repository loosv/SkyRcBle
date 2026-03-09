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

using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.Advertisement;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Enumeration;


namespace BtExperiment
{
    class DiscoveredDevice
    {
        public ulong Mac;
        public string Name = "";
        public int Rssi;
        public int RxCtr;
        public List<(byte DataType, byte[] Data)> DataSections = new List<(byte DataType, byte[] Data)>();
    }

    class Program
    {
        const int SCAN_TIME_S = 20;
        const string GapDevName = "Charger";
        static ulong macThatNeedsScan = 0;
        static bool macFoundInScan = false;

        static BluetoothLEAdvertisementWatcher _watcher;
        private static List<BleCharger> _devs = new List<BleCharger>();
        private static readonly System.Collections.Generic.Dictionary<ulong, DiscoveredDevice> advertisingDevs = new();
        static List<ulong> MacsRequested = new();

        static async Task ScanForDevices(int duration_s, int minRssi_dbm)
        {
            Console.WriteLine($"Starting BLE scan for {duration_s} seconds...");
            _watcher.Start();
            await Task.Delay(duration_s * 1000);
            _watcher.Stop();
            await Task.Delay(2000); // to wait for OnAdvertisementWatcherStopped()

            Console.WriteLine($"List of devices with RSSI >= {minRssi_dbm} dBm:");
            foreach (var entry in advertisingDevs)
            {
                DiscoveredDevice dev = entry.Value;
                if (dev.Rssi > minRssi_dbm)
                {
                    PrintAdvertisement(dev.Mac, dev.Rssi, dev.DataSections, dev.Name);
                }
            }
            Console.WriteLine($"List of chargers:");
            foreach (var entry in advertisingDevs)
            {
                DiscoveredDevice dev = entry.Value;
                // advertisement example: (1) 01 06; (2) 02 E0 FF; (22) FF 63 76 75 10 1C 64 31 30 30 30 38 33 01 00 00 00 00 01 19 16 00 A3;
                if (dev.DataSections.Count == 3)
                {
                    var section = dev.DataSections[2];
                    if (section.DataType == 0xFF ) // manufacturer data
                    {
                        byte[] dat = section.Data; // 0x63 is at index 0
                        ushort manufId = BitConverter.ToUInt16(dat, 0); // not a company ID. Just 16bit from the MAC address (observed with SkyRC MC3000, firmware version 1.21 and 1.25 of)
                        ushort macPart = (ushort)(dev.Mac & 0xFFFF);
                        if (manufId == macPart && dat.Length == 22 && dat[6] == 0x31 && dat[7] == 0x30 && dat[8] == 0x30 && dat[9] == 0x30 && dat[10] == 0x38 && dat[11] == 0x33)
                        {
                            PrintAdvertisement(dev.Mac, dev.Rssi, dev.DataSections, dev.Name);
                        }
                    }
                }
            }
        }


        // Connect to a known device (known MAC)
        // There are multiple cases
        //  if device is already connected: a new instance will be created (but "stale session" handling in discovery is necessary)
        //  if device is in system cache so the instance can be created immediately, the connected event follows either within seconds or hours or never
        //  if device is not in system cache anymore (~12h) a scan is necessary before connect.
        // Remark: when multiple devices are not in system cache (but are all in range), the scan puts all of them in system cache. 
        // Remark: if multiple devices are not in system cache and not in range, multiple scans will be prevented.
        static async Task<BleCharger?> ConnectToDeviceAndRunTask(ulong mac, int interval_ms, string logPath)
        {
            // Try to connect to a device already connected. See Quirk01 (stale sessions). Assuming no other application needs the device.
            string selector = BluetoothLEDevice.GetDeviceSelectorFromConnectionStatus(BluetoothConnectionStatus.Connected); // BLE only
            var infos = await DeviceInformation.FindAllAsync(selector);
            foreach (var info in infos)
            {
                BluetoothLEDevice? ble = null;
                try
                {
                    ble = await BluetoothLEDevice.FromIdAsync(info.Id);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Could not create instance from a already connected device: {ex.Message}");
                }

                // Console.WriteLine($" Name: {dev.Name}, ID: {dev.Id}, Kind: {dev.Kind}, Conn: {ble?.ConnectionStatus}, Paired: {dev.Pairing?.CanPair}");
                // dev.Name is DeviceName from GAP service

                string macStr = info.Id.Split('-')[1];
                ulong macCandidate = Convert.ToUInt64(macStr.Replace(":", ""), 16);
                if (mac == macCandidate)
                {
                    BleCharger dev2 = new BleCharger(mac, interval_ms, logPath);
                    Console.WriteLine($"{BleCharger.GetTs()} {mac:X12} already connected at start. GATT discovery may be slow (due to repeated AccessDenied error and due to 1000ms connection interval).");
                    bool success2 = await dev2.ConnectRequest();
                    if (success2)
                    {
                        _devs.Add(dev2);
                        return dev2;
                    }
                    else
                    {
                        Console.WriteLine($"{BleCharger.GetTs()} {mac:X12} Device with MAC found during scan but failed to connect.");
                        return null;
                    }
                }
            }

            BleCharger dev = new BleCharger(mac, interval_ms, logPath);
            bool success = await dev.ConnectRequest();
            if (success)
            {
                // The device is in system cache yet (may be out of range or disabled, so the ConnectionStatusChanged Listener may get the connection event hours later or never)
                _devs.Add(dev);
                return dev;
            }
            else // if the device is not in system cache anymore (after ~12h) a scan is necessary before connect.
            {
                if (advertisingDevs.Count > 0) // means device is not in range (otherwise it would be already in system cache during firs scan).
                {
                    Console.WriteLine($"{BleCharger.GetTs()} {mac:X12} Device is likely not in range. And because it is not in system cache it can not be added to list of devices (-> in this case no late connect is supported by the app). Skipping device.");
                    return null;
                }

                Console.WriteLine($"{BleCharger.GetTs()} {mac:X12} Direct connect failed. Windows needs to scan for it.");
                macFoundInScan = false;
                macThatNeedsScan = mac;
                _watcher.Start();
                for (int i = 0; i < SCAN_TIME_S; i++)
                {
                    if (macFoundInScan)
                        break;
                    await Task.Delay(1000);
                }
                _watcher.Stop();
                await Task.Delay(2000); // wait until the watcher is stopped // TODO: check if it might interfere with connect, maybe optimize time
                if (macFoundInScan)
                {
                    Console.WriteLine($"{BleCharger.GetTs()} {mac:X12} Device found during scan.");
                    bool success2 = await dev.ConnectRequest();
                    if (success2 == false)
                    {
                        Console.WriteLine($"{BleCharger.GetTs()} {mac:X12} Device found during scan, failed to connect.");
                        return null;
                    }
                    _devs.Add(dev);
                    return dev;
                }
                else
                {
                    Console.WriteLine($"{BleCharger.GetTs()} {mac:X12} Not found.");
                    return null;
                }
            }
        }

        // input e.g. "123456789ABC;123456789DEF"
        static List<ulong> GetMacs(string str)
        {
            List<ulong> macs = new List<ulong>();
            if (str.Length < 12)
                return macs;

            string[] parts = str.Split(';'); // if only 1 MAC is given, the array contains that MAC
            foreach (string part in parts)
            {
                if (part.Length == 12 && ulong.TryParse(part, System.Globalization.NumberStyles.HexNumber, null, out var mac))
                {
                    macs.Add(mac);
                }
                else
                {
                    Console.WriteLine($"Invalid MAC address or format not supported: {part}");
                }
            }
            return macs;
        }

        static async Task WaitTillUserAction()
        {
            while (true)
            {
                await Task.Delay(500);
                if (Console.KeyAvailable)
                {
                    ConsoleKeyInfo key = Console.ReadKey(intercept: true);
                    if (key.Key == ConsoleKey.Escape)
                    {
                        break;
                    }
                    else if (key.Key == ConsoleKey.R)
                    {
                        foreach (BleCharger dev in _devs)
                        {
                            dev.RestartCsv();
                        }
                    }

                }
            }
        }
        static async Task Cleanup()
        {
            Console.WriteLine("Disconnecting devices...");
            foreach (BleCharger dev in _devs)
            {
                await dev.Disconnect();
            }
            Console.WriteLine("Done.");
            await Task.Delay(1000);
        }

        const string guide =
@"Usage:
  prg [command] [MacAddr] [parameters]
Examples:\n\
  prg -l 20 -80   lists all advertising devices and the identified charger. 60s scan duration. Minimum RSSI -80dBm.
  prg -v C4E3046EB278;C4E3046EB279 <logPathAndFileName>  live view and log multiple chargers.
logPathAndFileName can be a full path with full file name or just a measurment name to which the MacAddress and file ending will be appended.";

        static async Task Main(string[] args)
        {
            AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
            _watcher = new BluetoothLEAdvertisementWatcher { ScanningMode = BluetoothLEScanningMode.Passive }; // Passive, otherwise heavy interference with Bluetooth Classic devices
            _watcher.Received += OnAdvertisementReceived;
            _watcher.Stopped += OnAdvertisementWatcherStopped;

            // guide
            if (args.Length < 2)
            {
                Console.WriteLine(guide);
                return;
            }
            Console.WriteLine("Press ESC or CTRL+C to stop sampling and disconnect.");


            if (args[0] == "-l")
            {
                if (args.Length == 3)
                {
                    Int32 scanDuration = 0;
                    Int32 minRssi = 0;
                    try
                    {
                        scanDuration = Convert.ToInt32(args[1]);
                        minRssi = Convert.ToInt32(args[2]);
                        await ScanForDevices(scanDuration, minRssi);
                        return;
                    }
                    catch (Exception) { }
                }
                Console.WriteLine("Parameter error.");
                return;
            }
            if (args[0] == "-v")
            {
                if (args.Length == 3)
                {
                    string logPath = "";
                    try
                    {
                        MacsRequested = GetMacs(args[1]);
                        logPath = args[2];
                        foreach (ulong mac in MacsRequested)
                        {
                            await ConnectToDeviceAndRunTask(mac, 0, logPath);
                            await Task.Delay(10000); // to avoid parallel discovery on mulitple devices.
                        }
                        await WaitTillUserAction();
                        await Cleanup();
                        return;
                    }
                    catch (Exception) { }
                }
                Console.WriteLine("Parameter error.");
                return;
            }
            await Cleanup();
        }

        // Analyze only the first received advertisement paket from each device. Add an entry in DeviceList (MAC and a free defined name). 
        private static void OnAdvertisementReceived(BluetoothLEAdvertisementWatcher watcher, BluetoothLEAdvertisementReceivedEventArgs args)
        {
            if (macThatNeedsScan != 0) // use case: if windows can not connect to a distinct device -> scan is required.
            {
                if (args.BluetoothAddress == macThatNeedsScan)
                {
                    macFoundInScan = true; // connect function waits for this.
                }
            }

            if (!advertisingDevs.ContainsKey(args.BluetoothAddress)) 
            {
                var dev = new DiscoveredDevice { Mac = args.BluetoothAddress, Name = args.Advertisement.LocalName, Rssi = args.RawSignalStrengthInDBm, RxCtr = 1 };
                foreach (var section in args.Advertisement.DataSections)
                {
                    byte[] data = section.Data.ToArray();
                    dev.DataSections.Add((section.DataType, data));
                }
                PrintAdvertisement(dev.Mac, dev.Rssi, dev.DataSections, dev.Name);
                advertisingDevs[args.BluetoothAddress] = dev;
            }
        }

        public static void PrintAdvertisement(ulong mac, int rssi, List<(byte DataType, byte[] Data)> dataSections, string deviceName)
        {
            Console.Write($"{mac:X12}; {rssi} dBm; ");
            foreach (var section in dataSections)
            {
                byte[] data = section.Data.ToArray();
                string hex = BitConverter.ToString(data).Replace("-", " ");
                Console.Write($"({data.Length}) {section.DataType:X2} {hex}; ");
            }
            if (deviceName.Length > 0)
                Console.WriteLine($"\"{deviceName}\""); // Many devices do not have a name in advertisement packets.
            else
                Console.WriteLine("");
        }

        private static void OnAdvertisementWatcherStopped(BluetoothLEAdvertisementWatcher watcher, BluetoothLEAdvertisementWatcherStoppedEventArgs args)
        {
            Console.WriteLine("Scan stopped.");
        }

        private static void OnProcessExit(object? sender, EventArgs e)
        {
            foreach (BleCharger dev in _devs)
            {
                dev.Disconnect().Wait();
            }
        }
    }
}
