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

// credits to https://github.com/kolinger/skyrc-mc3000

using System.Globalization;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Storage.Streams;

namespace BtExperiment
{
    // Writes the BATTERY_INFO measurements to CSV files. 1 CSV file per slot. Writes to same file as long as the application is running. No new file on stop/start or battery replacement. For new files press R.
    // Can start programs in a primitive way, see ChargerExperimentOnStartup()
    // Logging to C:\Users\user\AppData\Local\BleCharger\MAC_Timestamp_Slot.csv

    // Communication concept of the charger:
    //   The charger uses an unofficial, widely used "serial port profile" with 20byte characteristic.
    //   Only 1 characteristic with fixed size of 20 bytes is used for requests. Receiving responses by NOTIFICATION on same characteristic.
    //   Wait for response before the next request (this application just uses fixed delays).
    //   Requests are usually 20byte, fixed size (except CHARGER_WRITE_PROGRAM, which is sent as 2 packets, 20 byte each).
    //   Responses (by NOTIFICATION) are usually a single packet 20byte fixed size (except CHARGER_BATTERY_CHART)
    // BLE can get unavailable after ESD. Powercycle the Charger.

    // This implementation is fitted to BluetoothLEDevice class from Windows.Devices.Bluetooth
    // Tested under Windows 10 & 11 and .NET8. Should work starting with .NET5. Maybe lower "Target OS Version" must be set for some Windows installations.

    // BleCharger class
    // - ConnectRequest() creates a new BluetoothLEDevice instance (using MAC address) and triggers a new connection (by setting MaintainConnection=true)
    // - Discover() gets the "serial port" charactersistic and enables notifications for it. And stores the characteristic reference. If MaintainConnection=false, Discover() would initiate the connection.
    // - OnNotificationReceived handler: prints the received data and writes CSV files.
    // - ConnectionStatusChanged handler to rediscover the characteristic and re-enable notification.
    // - out of scope: scan for devices and device type filter.

    // Implemented connection sequences:
    //   If device is already connected:  ConnectRequest() -> Discover()
    //   If device is NOT connnected:     ConnectRequest();   ... OnConnectionStatusChanged()-> Discover()

    // General BLE properties under Windows:
    // - Can NOT directly connect to Devices never seen during a scan.
    // - BLE scan can impact other connections (during scan).
    // - Characteristic objects get invalid on connection loss.
    // - Use MaintainConnection=true instead of connection retries. Keep BluetoothLEDevice object. Still, get new characteristic objects on reconnect.
    // - if MaintainConnection==false, Uncached access is necessary to initiate a connection, e.g. by GetGattServicesAsync(Uncached).
    // - Do not unsubscribe listeners on connection loss. Only on program end: disable notifications on clients, unsubscribe listeners, dispose BLE objects.
    // - Characteristic Cache: Windows maintains a system-wide cache for attribute discovery (services/characteristics) and values.
    //   - According to documentation, entries in the cache become invalid ONLY when the remote device indicates a service change, or when the device becomes unpaired.
    //     - The peripheral sends a Service Changed indication/notification (via the 0x2A05 characteristic in the Generic Attribute service 0x1801). Charger does not seem to use it.
    //     - The device is unpaired (removed from paired devices). Charger does not use pairing.
    // - Attribute handles: Windows BLE stack does not allow addressing by attribute handles. But GattCharacteristic.AttributeHandle property exists.
    //     It returns the 16-bit handle assigned by the peripheral. But it can NOT be used to bypass discovery or recreate characteristics.
    // - Only 1 call to GetGattServicesAsync or GetCharacteristicsAsync() is possible reliably. Subsequent calls may fail for a while with "AccessDenied". There can be stale sessions from a previous run.
    // - Debugging: if a Bluetooth mouse freezes during debugging of connect procedure, step over or use a USB mouse.
    // - No connection is required or attempted for cached discovery. But usually Windows does connect on GetGattServicesAsync()

    enum BatteryType : byte
    {
        LiIon = 0,
        LiFe = 1,
        LiIion435 = 2,
        NiMH = 3,
        NiCd = 4,
        NiZn = 5,
        Eneloop = 6,
        Ram = 7,
        LTO = 8,
        NaIon = 9
    }
    enum Operation : byte
    {
        Charge = 0,
        Refresh = 1,
        StoreBreakin = 2,
        Discharge = 3,
        Cycle = 4
    }

    enum CycleMode : byte
    {
        CD = 0,
        CDC = 1,
        DC = 2,
        DCD = 3,
    }

    // SkyRC MC3000 program
    // Same structure for all battery types and programs (SkyRC MC3000, firmware 1.25, hw 2.20. As used by Android App 4.1.2).
    class ChargerProgram
    {
        public BatteryType battery_type; // 0=LiIon, 1=LiFe, 2=LiIon435, 3=NiMh, 4=NiCd, 5=NiZn, 6=Eneloop, 7=Ram, 8?=Lto, 9=NaIon
        public Operation operation_mode; // 0=charge, 1=refresh*, 2=store_breakin*, 3=discharge, 4=cycle
                                               // 1 refresh: LiIon_LiFe_LiIon435=CDC; NiMh_NiCd_Enelop=CDC_specificCurrent; NiZn_Ram=?
                                               // 2 store_breakin: LiIon_LiFe_LiIon435=store, NiMh_NiCd_Enelop=breakin, NiZn_Ram=not_available
        public ushort capacity;     // [mAh] 0=disabled, 1..50000: limit for charging AND discharging (if cylcing: for both), BleApp makes 100mAh steps, but it works with 1mAh steps.
        public ushort charge_current; // [mA]   50-3000  LiIon, NiMH.
        public ushort discharge_current; // [mA] 50-2000  LiIon, NiMH (lower current accepted).
        public ushort charge_end_voltage; // [mV]  1400..4400, battery type dependent.  Also the voltage limit for "store" operation (even if discharging).
        public ushort discharge_cut_voltage; // [mV]  500..3750, battery type dependent (see manual).
        public ushort charge_end_current; // [mA]    0..(charge_current-1), charge_current=disabled
        public ushort discharge_reduce_current; // [mA] 0..(discharge_current-1), discharge_current=disabled
        public byte charge_resting_time; // [minutes] 0..240
        public byte number_cycle; // 1-99  (BleApp sets 1 for charge or discharge too. But 0 works too).
        public CycleMode cycle_mode; // 0 = CD, 1 = CDC, 2 = DC, 3 = DCD
        public byte peak_sense_voltage; // [mV] negative sign implicit. Ranges 0=disabled,1-20. Default: 3 for all batterie types using it (NiMH, NiCd, Eneloop).
        public byte trickle_current; // [10mV] 0=disabled, 1-30   For NiMH, NiCd, Eneloop only. Default: 10 (for 100mA). If enalbed, restart_voltage should be 0.
        public ushort restart_voltage; // [mv]    0=disabled, 1300..4330, liion: 3980..4180;  if enabled, trickle_current should be 0
        public byte cut_temperature; // in [°C] or [°F]. Units defined by global setting. 0=disabled, 20-70 °C or 68-158 °F, default 45°C.
        public ushort cut_time; // [minutes] 0=disabled, 1..1440
        public byte discharge_resting_time; // [minutes] 0..240

        // restart_voltage vs trickle_current: charger allows only one of them to be set. App sets both (not checked if charger will do both).
        // trickle_time: [OFF/end/rest] is not available in BleApp. Available on charger and USB. For NiMH refresh all 3. For NimH charging only OFF/end.
        // temperature_unit: not available in program (not like USB)
        // Eneloop "MODEL" has no own field, it changes capacity and current limits.
        // Variable sizes: as transmitted over BLE and USB. For ranges and necessary combinations check Charger or BleApp.
        //
        // Names from manual / Name on charger / variable names
        //   "Charge Voltage max." / "TARGET VOLT" /  charge_end_voltage
        //   "Storage voltage" / "TARGET VOLT"  /  charge_end_voltage
        //   "Discharge voltage min." / "CUT VOLT" /  discharge_cut_voltage
        //   "Restart voltage" / "RESTART VOLT"  /  restart_voltage
        //
        // In "store" mode the charger automatically chages or discharges to "charge_end_voltage". discharge_cut_voltage is not used.
        //   charge_current and charge_end_current will be used, if battery needs charging.
        //   discharge_current and discharge_reduce_current used, if batter needs discharging.
        //   If "store" selected on Charger: charge_end_current=0, discharge_reduce_current=0. Is slow because waiting until the current is almost 0. So SoC is almost the same for charging and discharging path.
        //   If "store" selected in BleApp: charge_end_current=100, discharge_reduce_current=discharge_current. Is fast, but less precise (much bigger difference in SoC between charging and discharging path). 
        //
        // BleApp sets peak_sense_voltage & trickle_current only for NiMH+NiCd+Eneloop charging+cylcing (which makes sense).
        //
        // The charger silently ignores many (all?) unnecessary parameters and accept some wider ranges than shown on charger or BleApp.
        //
        // Some unnecessary stuff:
        //  BleApp always sets 2x current limits and 4x voltage limits
        //  BleApp sets sometimes restart_voltage unnecessarily.
        //  BleApp sets charge_end_current=100 for NiMH.
        //  BleApp sets restart_voltage for cycling (NiMh,LiIon)
        //  BleApp shows restart_voltage for NaIion, which is higher than charge_end_votlage
        //
        // BleApp 4.1.2 bugs:
        //  it probably does not get the system settings at all (with 0x61) because it sents the request without waiting for previous request version request (0x57)
    }

    class BleCharger
    {
        // Sky RC3000 charger UUIDs
        static readonly Guid chargerServiceUuid = new Guid("0000ffe0-0000-1000-8000-00805f9b34fb");
        static readonly Guid chargerChUuid = new Guid("0000ffe1-0000-1000-8000-00805f9b34fb"); // the unofficial, widely used "serial port profile" with 20byte data

        public BluetoothLEDevice? device;      // does NOT get invalid on connection loss. survives reconnect
        GattSession? session;                  // does NOT get invalid on connection loss.
        GattDeviceService? usedService;        // invalid on connection loss
        public GattCharacteristic? chCharger;  // invalid on connection loss
        Task? pollingTask = null;              // reconnect behavior: it will on-the-fly take the new chCharger after reconnect

        public ulong Mac; // e.g. 0x0000123456789ABC for "12:34:56:78:9A:BC"
        CancellationTokenSource cts;
        static System.Globalization.CultureInfo culture = CultureInfo.InvariantCulture;
        public string path = "";
        public string csvStartTime = "";
        public string temperatureUnit = "?"; // filled right after connect (by reading global settings from charger)

        // BleApp sets some unused values. Reasons unknown.
        // Charger accepts these profiles here, with zeroes in all unused values, and works as expected.
        // cut_temperature  for units see temperatureUnit
        ChargerProgram PrgLiIon_charge = new ChargerProgram
        {
            battery_type = BatteryType.LiIon,
            operation_mode = Operation.Charge,
            capacity = 4000,
            charge_current = 500,
            charge_end_voltage = 4200,
            charge_end_current = 50,
            cut_temperature = 45,
            // All undefined values are zero-initialized.
        };
        ChargerProgram PrgLiIon_discharge = new ChargerProgram
        {
            battery_type = BatteryType.LiIon,
            operation_mode = Operation.Discharge,
            discharge_current = 500,
            discharge_cut_voltage = 3000,
            discharge_reduce_current = 100,
            cut_temperature = 45, // °C
        };
        ChargerProgram PrgLiIon_store = new ChargerProgram
        {
            battery_type = BatteryType.LiIon,
            operation_mode = Operation.StoreBreakin,
            capacity = 4000,
            charge_current = 500,  // used if battery voltage is lower.
            discharge_current = 1000, // used if battery voltage is higher.
            charge_end_voltage = 3700,  // only this voltage limit is used for store.
            discharge_cut_voltage = 0, // ignored!
            charge_end_current = 0,   // used if battery voltage is lower. 0 means until ~5mA reached.
            discharge_reduce_current = 0, // used if battery voltage is higher. 0 means until ~5mA reached.
            cut_temperature = 45,
            cut_time = 600,
            // set charge_end_voltage to 3400 for <30% SoC with NCA and to 3500 for <30% with NMC
        };
        ChargerProgram PrgLiIon_cycle = new ChargerProgram
        {
            battery_type = BatteryType.LiIon,
            operation_mode = Operation.Cycle,
            capacity = 4000,
            charge_current = 900,
            discharge_current = 900,
            charge_end_voltage = 4200,
            discharge_cut_voltage = 3000,
            charge_end_current = 50,
            discharge_reduce_current = 100, // to disable reduction: set it =discharge_current
            charge_resting_time = 20,
            restart_voltage = 0, // disabled
            cut_temperature = 45,
            cut_time = 0, // disabled
            cycle_mode = CycleMode.DC,
            discharge_resting_time = 20,
            number_cycle = 2,
        };
        ChargerProgram PrgLiFe_charge = new ChargerProgram
        {
            battery_type = BatteryType.LiFe,
            operation_mode = Operation.Charge,
            capacity = 4000,
            charge_current = 500,
            charge_end_voltage = 3600,
            charge_end_current = 100,
            cut_temperature = 45,
        };
        ChargerProgram PrgNiMH_charge = new ChargerProgram // same for NiMH & Eneloop.
        {
            battery_type = BatteryType.NiMH,
            operation_mode = Operation.Charge,
            capacity = 5000, // adjust
            charge_current = 500, // adjust
            charge_end_voltage = 1650,
            peak_sense_voltage = 3,
            trickle_current = 0, // disabled
            restart_voltage = 0, // disabled
            cut_temperature = 45,
        };
        ChargerProgram PrgNiMH_discharge = new ChargerProgram
        {
            battery_type = BatteryType.NiMH,
            operation_mode = Operation.Discharge,
            capacity = 0,
            discharge_current = 50, 
            discharge_cut_voltage = 1000,
            discharge_reduce_current = 0,
            cut_temperature = 45,
        };
        ChargerProgram PrgNiMH_breakin = new ChargerProgram
        {
            // As sent by BleApp for a 2000mAh battery. Not all are shown in BleApp.
            // did not completely run this profile, so behavior details not verified.
            battery_type = BatteryType.NiMH,
            operation_mode = Operation.StoreBreakin, // break-in
            capacity = 2000,
            charge_current = 200, // derived from capacity
            discharge_current = 400,  // derived from capacity
            charge_end_voltage = 1650, // ignored
            charge_end_current = 100,  // maybe ignored (not visible on Charger or BleApp), didnt test if it is reduced
            discharge_cut_voltage = 1000,
            discharge_reduce_current = 500, // ignored (reduction can not be enabled on charger)
            number_cycle = 1,  // ignored
            cycle_mode = CycleMode.CDC, // CDC or DCD can be selected.
            peak_sense_voltage = 0, // ignored  (can not be enabled on charger)
            trickle_current = 0, // ignored  (can not be enabled on charger)
            charge_resting_time = 60, 
            discharge_resting_time = 60,
            restart_voltage = 1350, // ignored
            cut_temperature = 45,
            // "ignored" = setting is not changeable on the charger.
        };


        public static string getTs()
        {
            //string ts = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"); // the most simple ISO 8601 format
            string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"); // compatible with LibreOffice Calc (just selecting "YMD" will parse the whole timestamp including time)
            return ts;
        }
        ushort SwapEndian16(ushort value)
        {
            return (ushort)((value >> 8) | (value << 8));
        }
        public static byte ToCelsius(byte fahrenheit) // input range 32°F - 255°F
        {
            return (byte)Math.Round((fahrenheit - 32) * 5.0 / 9.0);
        }
        public static byte ToFahrenheit(byte celsius) // input range 0°C - 123°C
        {
            return (byte)Math.Round(celsius * 9.0 / 5.0 + 32.0);
        }
        void InitChargerCsv()
        {
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            //string folder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            path = Path.Combine(folder, "BleCharger");
            Console.WriteLine("CSV path: " + path);

            csvStartTime = DateTime.Now.ToString("yyyyMMddTHHmmss"); //  compact ISO 8601 format without separators, with T between date and time 
        }
        public void RestartCsv()
        {
            Console.WriteLine("New CSV files will be created on next sample.");
            csvStartTime = DateTime.Now.ToString("yyyyMMddTHHmmss"); // will trigger new file in WriteChargerCsv(), see pathComplete
            // remark: this is not syncrhonized to all slots (so some slots may already be sampled in the old CSV files).
        }

        void WriteChargerCsv(ulong mac, byte channel, short voltage, short current, short capacity, short temperature, string status)
        {
            if (path == null || path == "")
                return;

            string macStr = mac.ToString("X12");
            string pathComplete = Path.Combine(path, $"{macStr}_{csvStartTime}_{channel}.csv");
            const string separator = ";";

            StringWriter writer = new StringWriter();
            if (pathComplete == null)
                return;

            if (!File.Exists(pathComplete))
            {
                if (Path.IsPathFullyQualified(path))
                {
                    Directory.CreateDirectory(path);
                }

                // create titles
                writer.Write($"timestamp;voltage [mV];current [mA];capacity [mAh];temperature [°{temperatureUnit}];state\n");
                File.WriteAllText(pathComplete, writer.ToString());
            }

            //string line = GenerateIso8601_LocalTime()+ separator; // import in LibreOffice fails
            string line = getTs() + separator; 
            // write CSV line
            line += voltage.ToString() + separator;
            line += current.ToString() + separator;
            line += capacity.ToString() + separator;
            line += temperature.ToString() + separator;
            line += status;
            line += "\n";
            File.AppendAllText(pathComplete, line);
        }

        public BleCharger(ulong Mac)
        {
            this.Mac = Mac;
            cts = new CancellationTokenSource();
            InitChargerCsv();
        }

        const byte CHARGER_HEADER = 0x0F;

        const byte CHARGER_WRITE_PROGRAM = 0x11; // accepted only with 40byte request (split in 2*20byte packets)
        const byte CHARGER_START_PROGRAM = 0x05;
        const byte CHARGER_STOP_PROGRAM = 0xFE;
        const byte CHARGER_BATTERY_INFO = 0x55;
        const byte CHARGER_BATTERY_CHART = 0x56;
        const byte CHARGER_READ_VERSIONS = 0x57;
        const byte CHARGER_READ_SYSTEM_SETTINGS = 0x61;
        const byte CHARGER_WRITE_SYSTEM_SETTINGS = 0x63;
        const byte CHARGER_WRITE_UNKNOWN_A = 0x65;
        const byte CHARGER_WRITE_UNKNOWN_B = 0x66;

        // no responses to other command codes (0x00 .. 0xFF tested)
        string[] CARGER_CELL_TYPES = { "LiIon", "LiFe", "LiIo4_35", "NiMH", "NiCd", "NiZn", "Eneloop", "Ram", "LTO", "NaIon"};
        private static readonly Dictionary<byte, string> CHARGER_STATUS = new Dictionary<byte, string>
        {   // values from https://github.com/kolinger/skyrc-mc3000, seems to be the same for BLE (0,1,2,4,132,133,134,137 verified)
            { 0,   "Standby" },
            { 1,   "Charge" },
            { 2,   "Discharge" },
            { 3,   "Pause" },
            { 4,   "Completed" },
            { 128, "Input low voltage" },
            { 129, "Input high voltage" },
            { 130, "ADC MCP3424-1 error" },
            { 131, "ADC MCP3424-2 error" },
            { 132, "Connection break" },
            { 133, "Check voltage" },
            { 134, "Capacity limit reached" },
            { 135, "Time limit reached" },
            { 136, "SysTemp too hot" },
            { 137, "Battery too hot" },
            { 138, "Short circuit" },
            { 139, "Wrong polarity" },
            { 140, "Bad battery (high IR)" }
        };
        public string GetStatusDescription(byte statusCode)
        {
            return CHARGER_STATUS.TryGetValue(statusCode, out var description) ? description : $"Stat_{statusCode}";
        }
        byte CalculateChecksum(byte[] data, int len)
        {
            byte sum = 0;
            for (int i = 0; i < len; i++)
                sum += data[i]; // overflow ok
            return sum;
        }

        public async Task ChargerWrite(GattCharacteristic? ch, byte[] data)
        {
            if (ch == null)
                return;
            IBuffer buffer = data.AsBuffer();
            try
            {
                await ch?.WriteValueAsync(buffer);
            }
            catch (ObjectDisposedException ) {}
        }

        public async Task ChargerBatteryInfo(byte slot) // slot 0,1,2,3
        {
            if (chCharger == null)
                return;
            byte[] req = new byte[20]; // zeroinitialized
            req[0] = CHARGER_HEADER;
            req[1] = CHARGER_BATTERY_INFO;
            req[2] = slot;
            req[19] = CalculateChecksum(req, 19);
            await ChargerWrite(chCharger, req);
            // response example: 0F 55 00 00 03 00 02 00 0A 0F A3 01 8A 00 00 19 00 B5 01 7F
        }

        public async Task ChargerGetVoltageChart(byte slot) // slot 0,1,2,3
        {
            if (chCharger == null)
                return;
            byte[] req = new byte[20]; // zeroinitialized
            req[0] = CHARGER_HEADER;
            req[1] = CHARGER_BATTERY_CHART;
            req[2] = slot;
            req[19] = CalculateChecksum(req, 19);
            await ChargerWrite(chCharger, req);
            // response: 13 notifications with 120 voltage values (byte0 = 0x0F, byte1 = slot, byte2&3 = sampling_interval_s (1,2,4,8,16,32,64,..?), byte4-243 = voltages, byte244 = checksum).
            // it changes the sampling interval at 120s, at 240s, at 480s, at 960s, at 1920s ...
        }

        public async Task ChargerStartProgram(byte slots)
        {
            if (chCharger == null)
                return;
            byte[] req = new byte[20];
            req[0] = CHARGER_HEADER;
            req[1] = CHARGER_START_PROGRAM;
            req[2] = slots;
            req[19] = CalculateChecksum(req, 19);
            await ChargerWrite(chCharger, req);
            // response example: 0F 05 0F F0 FF FF 00 00 00 00 00 00 00 00 00 00 00 00 00 11
        }

        // It stops any running programs (remotely started and locally started on charger).
        // It does not modify slot memory, so any program (also locally started) can be restarted.
        public async Task ChargerStopProgram(byte slots)
        {
            if (chCharger == null)
                return;
            byte[] req = new byte[20];
            req[0] = CHARGER_HEADER;
            req[1] = CHARGER_STOP_PROGRAM;
            req[2] = slots;
            req[19] = CalculateChecksum(req, 19);
            await ChargerWrite(chCharger, req);
            // response example: 0F FE 0F F0 FF FF 00 00 00 00 00 00 00 00 00 00 00 00 00 0A 
        }

        //  BleApp sends it as first packet.
        public async Task ChargerGetVersions(ulong mac)
        {
            if (chCharger == null)
                return;
            byte[] req = new byte[20];
            req[0] = CHARGER_HEADER;
            req[1] = CHARGER_READ_VERSIONS;
            req[2] = 0x00;
            req[3] = (byte)(mac); // BleApp adds MAC here. sending 6 zeros leads to same response.
            req[4] = (byte)(mac >> 8);
            req[5] = (byte)(mac >> 16);
            req[6] = (byte)(mac >> 24);
            req[7] = (byte)(mac >> 32);
            req[8] = (byte)(mac >> 40);
            req[19] = CalculateChecksum(req, 19);
            await ChargerWrite(chCharger, req);
            // response example: 0F-57-00-31-30-30-30-38-33-01-00-00-00-00-01-19-16-00-A3-29
            // byte 3-8 are: ASCII "100083", unknown meaning (same on 2 SkyRC MC3000, same for firmware 1.21 and 1.25)
            // byte 14-15: 0x0119 = 1.25 Software version
            // byte 16: 0x16 = 2.2 Hardware version
            // byte 17-18: unknown
            // byte 19: not a checksum, changes often
        }

        public async Task ChargerGetSettings()
        {
            if (chCharger == null)
                return;
            byte[] req = new byte[20];
            req[0] = CHARGER_HEADER;
            req[1] = CHARGER_READ_SYSTEM_SETTINGS; 
            req[19] = CalculateChecksum(req, 19);
            await ChargerWrite(chCharger, req);
            // response example: 0F-61-00-00-01-01-09-2A-F8-00-00-00-00-00-00-00-00-00-00-9D
        }
        public async Task ChargerWriteSettings()
        {
            const ushort minInputVoltage = 11000; // mV, range 10000 ... 12000

            if (chCharger == null)
                return;
            byte[] req = new byte[20];
            req[0] = CHARGER_HEADER;
            req[1] = CHARGER_WRITE_SYSTEM_SETTINGS;
            req[2] = 0x00; // 0 = °C, 1 = °F
            req[3] = 0x00; // 0 = system beep off, 1 = system beep on
            req[4] = 0x01; // 0 = backlight off, 1 = backlight auto, 2 = 1min, 3=3min, 4=5min, 5=always on
            req[5] = 0x00; // 0 = screen saver off, 1 = screen saver on (+ beeps, when system beep is on)
            req[6] = 0x00; // 0 = fan auto, 1 = fan off, 2 = fan ON, 3..9 = 20°C .. 50°C in 5°C steps
            req[7] = (byte)(minInputVoltage >> 8);
            req[8] = (byte)(minInputVoltage & 0xFF);

            req[19] = CalculateChecksum(req, 19);
            await ChargerWrite(chCharger, req);
        }

        // For probing other commands (only 9 responding)
        public async Task ChargerExperiment1(byte cmd)
        {
            byte[] req = new byte[20];
            req[0] = CHARGER_HEADER;
            req[1] = cmd; 
            //req[2] = 1;
            req[19] = CalculateChecksum(req, 19);
            await ChargerWrite(chCharger, req);
        }

        // For probing other commands (only 1 responding)
        public async Task ChargerExperiment2(byte cmd)
        {
            byte[] req = new byte[40];
            req[0] = CHARGER_HEADER;
            req[1] = cmd;
            //req[2] = 1;
            req[39] = CalculateChecksum(req, 39);
            byte[] part1 = req.AsSpan(0, 20).ToArray();
            byte[] part2 = req.AsSpan(20, 20).ToArray();
            await ChargerWrite(chCharger, part1);
            await ChargerWrite(chCharger, part2);
        }

        // Precondition: stop the currently running programs with ChargerStopProgram(), otherwise the write will be ignored. No other commands required before.
        // After programming, wait ~100ms and call ChargerStartProgram().
        // BLE app does:  stop, wait 200ms, write program, wait 30ms, start.
        // Remark: Program will be loaded directly into slot memory, like one of the programs selectable on charger. None of the programs are overwritten.
        // If stopped, it can NOT be reactivated by button click on the charger (the currently shown program will be started).
        // Recomendation: create programs which are inherently safe for the connected battery. Do not think the PC will be able to abort a unsafe program in time.
        public async Task ChargerSendSlotProgram(byte slots, ChargerProgram p) // slot = 0x01,0x02,0x04,0x08 for slots 0,1,2,3; 0x0F for all slots  (bitmask)
        {
            // Naming same as https://github.com/kolinger/skyrc-mc3000 (for USB).
            // But the order of values is different for BLE (compared to USB).
            // Most values do have same size, resolution and range, except: trickle_current has different resolution (10mV), no trickle_time and no temperature_unit over BLE

            // But compared to USB, the structure is slightly reorganized. Checksum includes first 2 bytes. No 0xFFFF at the end.
            // The structure is the same for all battery programs.
            // For meaning and ranges see ChargerProgram class
            if (chCharger == null)
                return;

            byte[] req = new byte[40]; // zeroinitialized
            req[0] = CHARGER_HEADER;
            req[1] = CHARGER_WRITE_PROGRAM;
            req[2] = slots;
            req[3] = (byte)p.battery_type;
            req[4] = (byte)p.operation_mode;
            req[5] = (byte)(p.capacity >> 8);  // big endian
            req[6] = (byte)(p.capacity & 0xFF);
            req[7] = (byte)(p.charge_current >> 8);
            req[8] = (byte)(p.charge_current & 0xFF);
            req[9] = (byte)(p.discharge_current >> 8);
            req[10] = (byte)(p.discharge_current & 0xFF);
            req[11] = (byte)(p.charge_end_voltage >> 8);
            req[12] = (byte)(p.charge_end_voltage & 0xFF);
            req[13] = (byte)(p.discharge_cut_voltage >> 8);
            req[14] = (byte)(p.discharge_cut_voltage & 0xFF);
            req[15] = (byte)(p.charge_end_current >> 8);
            req[16] = (byte)(p.charge_end_current & 0xFF);
            req[17] = (byte)(p.discharge_reduce_current >> 8);
            req[18] = (byte)(p.discharge_reduce_current & 0xFF);
            req[19] = p.charge_resting_time;
            req[20] = p.number_cycle;
            req[21] = (byte)p.cycle_mode;
            req[22] = p.peak_sense_voltage;
            req[23] = p.trickle_current;
            req[24] = (byte)(p.restart_voltage >> 8);
            req[25] = (byte)(p.restart_voltage & 0xFF);
            req[26] = p.cut_temperature; // in [°C] or [F], depends on system setting.
            req[27] = (byte)(p.cut_time >> 8);
            req[28] = (byte)(p.cut_time & 0xFF);
            req[29] = p.discharge_resting_time;
            req[39] = CalculateChecksum(req, 39);

            Console.WriteLine($"{Mac:X12} PRG: {BitConverter.ToString(req)}");

            // Split in 20 byte chunks. Sending 40byte fails with: "Value does not fall within the expected range."
            byte[] part1 = req.AsSpan(0, 20).ToArray();
            byte[] part2 = req.AsSpan(20, 20).ToArray();
            await ChargerWrite(chCharger, part1); 
            await ChargerWrite(chCharger, part2);
            // No delay between chunks allowed (otherwise charger ignores both chunks, e.g.: 100ms is already too long. BleApp sends the second half after 30ms (probably just in the next connection interval).
            // expected response, e.g.: 0F 11 0F F0 FF FF 00 00 00 00 00 00 00 00 00 00 00 00 00 1D    with byte2 = slots (0x0F here = all slots selected)
            // on successfull program completion the charger sends: 0F-B0-00-01-00-00-00-00-00-00-00-00-00-00-00-00-00-00-00-C0  (with byte2 = slot number)
        }

        public async Task ChargerStartProgram(byte slots, ChargerProgram p)
        {
            if (temperatureUnit == "C")
                p.cut_temperature = 45;
            else
                p.cut_temperature = ToFahrenheit(45);

            await Task.Delay(250); // guard
            await ChargerStopProgram(slots);
            await Task.Delay(250); // BLE app waits 200ms
            await ChargerSendSlotProgram(slots, p); // Ble app: response 12ms after request
            await Task.Delay(250); // BLE app waits for response (~30ms). Here replaced by fixed delay after request.
            await ChargerStartProgram(slots);
            await Task.Delay(250); // guard
        }


        public async Task ChargerExperimentOnStartup()
        {

            Console.WriteLine($"{Mac:X12} DevStatus: {device?.ConnectionStatus} ConnStatus: {session?.SessionStatus}.");

            await ChargerGetVersions(Mac);
            await Task.Delay(250);
            await ChargerGetSettings(); // Do not deactivate this call. It gets the temperature units from charger.
            await Task.Delay(250);

            //await ChargerWriteSettings();
            //await Task.Delay(250);

            //await ChargerStartProgram(0x01, PrgLiIon_charge);
        }

        public void OnNotificationReceived(GattCharacteristic sender, GattValueChangedEventArgs args)
        {
            // This callback is registerd only to ONE device. No need to extract the identity of the device from service.
            var data = new byte[args.CharacteristicValue.Length];
            DataReader.FromBuffer(args.CharacteristicValue).ReadBytes(data);
            string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

            if (sender.Uuid == chargerChUuid)
            {
                if (data[0] == CHARGER_HEADER && data[1] == CHARGER_BATTERY_INFO)
                {
                    // e.g. 0F-55-03-00-00-00-01-09-A8-0D-6D-00-64-00-44-18-00-01-0F-63,   big endian
                    byte slot = data[2];
                    byte cellType = data[3];
                    byte mode = data[4];
                    byte count = data[5];
                    byte status = data[6];
                    string statusStr = GetStatusDescription(status);
                    ushort time = SwapEndian16(BitConverter.ToUInt16(data, 7));
                    short voltage = (short)SwapEndian16(BitConverter.ToUInt16(data, 9));
                    short current = (short)SwapEndian16(BitConverter.ToUInt16(data, 11));
                    short capacity = (short)SwapEndian16(BitConverter.ToUInt16(data, 13));
                    byte temperature = data[15]; // in [°C] or [°F], depends on system setting
                    ushort resistance = SwapEndian16(BitConverter.ToUInt16(data, 16)); // 0, 1, 0xFFFF = not available
                    byte led = data[18];
                    byte checksum = data[19]; // ignore, BLE has a CRC

                    if (status != 0 || time != 0 || voltage != 0 )
                    {
                        Console.WriteLine($"{ts} {Mac:X12} {slot} {cellType} {mode} {count} {statusStr} - {time} s, {voltage} mV, {current} mA, {capacity} mAh, {temperature} °{temperatureUnit}");
                        WriteChargerCsv(Mac, slot, voltage, current, capacity, temperature, statusStr);
                    }
                }
                else if (data[0] == CHARGER_HEADER && data[1] == CHARGER_READ_SYSTEM_SETTINGS)
                {
                    if (data[2] == 0x00)
                    {
                        temperatureUnit = "C";
                    }
                    else
                    {
                        temperatureUnit = "F";
                    }
                    Console.WriteLine($"{ts} {Mac:X12} NOTIFY: {BitConverter.ToString(data)}"); // all other notifications
                }
                else
                {
                    Console.WriteLine($"{ts} {Mac:X12} NOTIFY: {BitConverter.ToString(data)}"); // all other notifications
                }
            }
            else
            {
                Console.WriteLine($"{ts} {Mac:X12} NOTIFY {sender.Uuid}: {BitConverter.ToString(data)}"); // robustness
            }
        }

        public async Task ChargerPolling()
        {
            await ChargerExperimentOnStartup();

            while (!cts.IsCancellationRequested)
            {
                try
                {
                    for (byte channelIndex = 0; channelIndex <= 3; channelIndex++)
                    {
                        await ChargerBatteryInfo(channelIndex);
                        await Task.Delay(250, cts.Token); // if shorter (e.g. 100ms), duplicate responses may occur (more responses than requests).
                        if (cts.IsCancellationRequested)
                            return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }

                await Task.Delay(9600, cts.Token);
            }
        }
        public async Task DisableNotificationAsync(GattCharacteristic? characteristic)
        {
            if (characteristic == null)
                return;
            try
            {
                var status = await characteristic.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue.None); // even if disconnected already, call this. It is fast.
                if (status == GattCommunicationStatus.Success)
                {
                    Console.WriteLine($"{Mac:X12} Notifications disabled on characteristic {characteristic.Uuid}");
                    characteristic.ValueChanged -= OnNotificationReceived;
                }
                else
                {
                    Console.WriteLine($"{Mac:X12} Failed to disable notifications: {status}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{Mac:X12} Exception while disabling notifications: {ex.Message}.");
            }
        }

        public async void OnConnectionStatusChanged(BluetoothLEDevice sender, object args)
        {
            // on Windows 10 & 11 the reconnection handling can be simplified by setting MaintainConnection == true (setting once, after creating the BluetoothLEDevice instance)
            string ts = getTs();
            if (sender.ConnectionStatus == BluetoothConnectionStatus.Disconnected)
            {
                Console.WriteLine($"{ts} {Mac:X12} disconnected.");
            }
            else if (sender.ConnectionStatus == BluetoothConnectionStatus.Connected)
            {
                Console.WriteLine($"{ts} {Mac:X12} connected. Discovering ...");
                try
                {
                    bool success = await Discover();
                    if (success == true && pollingTask == null)
                    {
                        pollingTask = ChargerPolling();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ts} {Mac:X12} Exception during reconnect handling: {ex.Message}");
                }
            }
        }

        // TODO: make a separate Create() function for the case with already connected device?
        public async Task ConnectRequest() // Can be called if device is disconnected OR connected.
        {
            Console.WriteLine($"{getTs()} {Mac:X12} creating BLE device and session ...");
            device = await BluetoothLEDevice.FromBluetoothAddressAsync(Mac);
            if (device == null)
            {
                Console.WriteLine($"{getTs()} {Mac:X12} Failed to create instance.");
                return;
            }
            session = await GattSession.FromDeviceIdAsync(device.BluetoothDeviceId);
            device.ConnectionStatusChanged += OnConnectionStatusChanged;
            session.MaintainConnection = true; // For automatic reconnect. Also triggers the first connection (otherwise GetGattServicesAsync(BluetoothCacheMode.Uncached) can be used to initiate a connection)

            if (device.ConnectionStatus == BluetoothConnectionStatus.Connected) // if already connected, there will be no Connected event, so call Discover() here
            { // ? (session.SessionStatus == GattSessionStatus.Active)
                Console.WriteLine($"{getTs()} {Mac:X12} already connected. Discovering ...");
                bool success = await Discover();
                if (success)
                {
                    pollingTask = ChargerPolling();
                }
            }
        }

        public async Task<bool> Discover()
        {
            // The 2 retry loops usually help after stale sessions to the BLE device (when the application was not closed with ESC before, so no Dispose was called).
            // This problem is windows specific (no hardware bug). The problem is, only 1 call of GetGattServicesAsync or GetCharacteristicsAsync() works reliably. Subsequent calls may fail for a while with "AccessDenied". Subsequent can mean: from a stale session.
            // TODO: add Cleanup of BLe objects to error paths 
            if (device == null || device.ConnectionStatus != BluetoothConnectionStatus.Connected)
            {
                Console.WriteLine($"{getTs()} {Mac:X12} BluetoothConnectionStatus is {device?.ConnectionStatus}. Abort.");
                return false;
            }

            // Just get reference to one characteristic and store it to chCharger
            try
            {
                GattDeviceServicesResult srvResult;
                srvResult = await device?.GetGattServicesAsync(BluetoothCacheMode.Uncached);
                int ctr1 = 0;
                while ((srvResult.Status != GattCommunicationStatus.Success || srvResult.Services.Count == 0) && ctr1 < 10) // usually AccessDenied error
                {
                    Console.WriteLine($"{getTs()} {Mac:X12} Waiting for GATT Services.");
                    ctr1++;
                    await Task.Delay(1000);
                    srvResult = await device?.GetGattServicesAsync(BluetoothCacheMode.Uncached);
                }
                if (srvResult.Status != GattCommunicationStatus.Success)
                {
                    Console.WriteLine($"{getTs()} {Mac:X12} Service discovery failed: {srvResult.Status}");
                    return false;
                }
                if (device == null || device.ConnectionStatus != BluetoothConnectionStatus.Connected)
                {
                    Console.WriteLine($"{getTs()} {Mac:X12} Device already lost.");
                    return false;
                }
                foreach (GattDeviceService srv in srvResult.Services)
                {
                    if (srv.Uuid == chargerServiceUuid)
                    {
                        var chResult = await srv.GetCharacteristicsAsync(BluetoothCacheMode.Uncached); // recommeded to not mix Cached/Uncached
                        int ctr2 = 0;
                        while ((chResult.Status != GattCommunicationStatus.Success || chResult.Characteristics.Count == 0) && ctr2 < 12) // usually AccessDenied error
                        {
                            Console.WriteLine($"{getTs()} {Mac:X12} Waiting for GATT Characteristics.");
                            ctr2++;
                            await Task.Delay(1000);
                            chResult = await srv.GetCharacteristicsAsync(BluetoothCacheMode.Uncached);
                        }
                        if (chResult.Status != GattCommunicationStatus.Success || chResult.Characteristics.Count == 0)
                        {
                            Console.WriteLine($"{getTs()} {Mac:X12} Characteristic discovery failed.");
                            return false;
                        }
                        else
                        {
                            chCharger = chResult.Characteristics.FirstOrDefault(c => c.Uuid == chargerChUuid);
                            if (chCharger == null)
                            {
                                Console.WriteLine($"{getTs()} {Mac:X12} Charger characteristic not found.");
                                return false;
                            }
                        }
                        usedService = srv;
                    }
                    else
                    {
                        srv.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{getTs()} {Mac:X12} First uncached access failed. For now the Discovery is abborted without any handling. Exception message: {ex.Message}");
                return false;
            }

            // try to read connection parameters 
            try
            {
                BluetoothLEConnectionParameters? par = device?.GetConnectionParameters(); // is not available on Win10. Availabl on Win11.
                Console.WriteLine($"{Mac:X6} Interval {par?.ConnectionInterval * 1.25:F0} ms, Timeout {par?.LinkTimeout / 100.0:F1} s, Latency {par?.ConnectionLatency}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{Mac:X6} GetConnectionParameters() is not available on this platform.");
            }


            // Enable notification on chCharger
            if (chCharger != null)
            {
                chCharger.ValueChanged -= OnNotificationReceived; // to prevent double-subscription, which would lead to multiple calls of the OnNotificationReceived()
                chCharger.ValueChanged += OnNotificationReceived;
                var cccdStatus = await chCharger.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue.Notify);
                if (cccdStatus == GattCommunicationStatus.Success)
                {
                    Console.WriteLine($"{getTs()} {Mac:X12} Notifications enabled");
                    return true;
                }
                else
                {
                    Console.WriteLine($"{getTs()} {Mac:X12} Failed to enable notifications");
                    chCharger.ValueChanged -= OnNotificationReceived;
                    chCharger = null;
                }
            }

            return false;
        }

        public async Task Disconnect()
        {
            Console.WriteLine($"{getTs()} {Mac:X12} Disconnecting ...");

            cts?.Cancel();
            await Task.Delay(100);

            if (session != null)
            {
                session.MaintainConnection = false;
                session.Dispose();
                session = null;
            }

            if (device != null)
            {

                await DisableNotificationAsync(chCharger); // will also work, if connection was already lost
                chCharger = null;  // reference to characteristic object
                usedService?.Dispose(); // helpful, at least for debugging, BLE connection is closed faster.
                usedService = null;

                device.ConnectionStatusChanged -= OnConnectionStatusChanged;
                device.Dispose();
                device = null;
            }
        }
    }
}

