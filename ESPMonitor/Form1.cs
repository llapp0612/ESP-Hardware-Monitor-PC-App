using ESPMonitor.Properties;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using OpenHardwareMonitor.Hardware;
using System.IO.Ports;
using System.Threading;
using System.Text.Json;


namespace ESPMonitor
{
    public partial class Form1 : Form
    {
        private NotifyIcon notifyIcon;
        private static Point _Location;

        private const string RunKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
        private const string AppName = "ESP Monitor";
        private static bool serialPortopened = false;
        private static bool showData = false;
        private static bool sendReset = false;

        // Statische Eigenschaft, um den Wert von numericUpDown1 zu speichern
        public static decimal NumericUpDownValue { get; set; }

        public Form1()
        {
            InitializeComponent();
            InitializeSystray();
            this.Resize += MainForm_Resize;
            // Deaktiviere das maximale Fenster und verhindere das manuelle Ändern der Größe
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            SerialPort serialPort = connectToPort(this);
            serialPort.DataReceived += SerialPort_DataReceived;
            StartDataTransmission(serialPort);
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                // Konvertieren Sie den Sender in ein SerialPort-Objekt
                SerialPort serialPort = (SerialPort)sender;
                // Überprüfen, ob Daten im Puffer verfügbar sind
                if (serialPort.BytesToRead > 0)
                {
                    // Lesen Sie die empfangenen Daten aus der seriellen Schnittstelle
                    string receivedData = serialPort.ReadLine();

                    // Geben Sie die empfangenen Daten in der Konsole aus
                    Console.WriteLine("Received data: " + receivedData);
                    // Zugriff auf die TextBox und Schreiben der empfangenen Daten
                    textBox1.Invoke((MethodInvoker)delegate
                    {
                        textBox1.AppendText("Received data: " + receivedData + Environment.NewLine);
                    });
                }
            }
            catch(Exception ex)
            {

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Überprüfen, ob die Anwendung mit Windows gestartet wird
            bool runWithWindows = IsAppSetToRunWithWindows();

            // CheckBox entsprechend aktualisieren
            checkBox1.Checked = runWithWindows;
            checkBox2.Checked = Settings.Default.MinimizeAtStartup;
            numericUpDown1.Value = Settings.Default.NumericUpDownValue;
            NumericUpDownValue = Settings.Default.NumericUpDownValue;

            // Überprüfen, ob die Anwendung minimiert gestartet werden soll
            if (Settings.Default.MinimizeAtStartup)
            {
                // Anwendung minimieren
                WindowState = FormWindowState.Minimized;
            }
            
            LoadWindowLocation();
            _Location = this.Location;
        }

        private bool IsAppSetToRunWithWindows()
        {
            // Code zum Überprüfen, ob die Anwendung mit Windows gestartet wird
            // Rückgabewert true, wenn die Anwendung mit Windows gestartet wird, sonst false
            return Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run").GetValue("ESP Monitor") != null;
        }

        private void InitializeSystray()
        {
            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = Resources.car_esp_icon_138809;// Setzen Sie das Symbol für das Systray
            notifyIcon.Text = "ESP Monitor"; // Text für das Systray-Symbol
            notifyIcon.Visible = true;

            // Ereignishandler für Doppelklick auf das Systray-Symbol
            notifyIcon.DoubleClick += NotifyIcon_DoubleClick;

            // Menü für das Systray-Symbol erstellen
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Restore", null, RestoreMenuItem_Click);
            contextMenu.Items.Add("Exit", null, ExitMenuItem_Click);

            // Menü mit dem Systray-Symbol verknüpfen
            notifyIcon.ContextMenuStrip = contextMenu;
        }

        private void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            // Anwendung wiederherstellen, wenn doppelt auf das Systray-Symbol geklickt wird
            Show();
            WindowState = FormWindowState.Normal;
            this.Visible = true;
            this.ShowInTaskbar = true;
        }

        private void RestoreMenuItem_Click(object sender, EventArgs e)
        {
            // Anwendung wiederherstellen, wenn im Kontextmenü auf "Restore" geklickt wird
            Show();
            WindowState = FormWindowState.Normal;
            SaveWindowLocation();
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            // Anwendung beenden, wenn im Kontextmenü auf "Exit" geklickt wird
            Close();
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            // Überprüfen, ob die Anwendung minimiert wurde, und in das Systray minimieren
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
                this.Visible = false;
                this.ShowInTaskbar = false;
                _Location = this.Location;
                SaveWindowLocation();
            }
        }

        private void SaveWindowLocation()
        {
            // Speichern der Fensterposition in den Anwendungseinstellungen
            Settings.Default.WindowState = this.WindowState;

            if (this.WindowState == FormWindowState.Normal)
            {
                Settings.Default.WindowLocation = this.Location;
                Settings.Default.Size = this.Size;
            }
            else
            {
                Settings.Default.WindowLocation = this.RestoreBounds.Location;
                Settings.Default.Size = this.RestoreBounds.Size;
            }

            Settings.Default.MinimizeAtStartup = checkBox2.Checked;

            Settings.Default.Save();
        }

        private void LoadWindowLocation()
        {
            // Laden der Fensterposition aus den Anwendungseinstellungen
            if (Settings.Default.WindowLocation != Point.Empty)
            {
                if (Settings.Default.WindowState == FormWindowState.Normal)
                {
                    this.Location = Settings.Default.WindowLocation;
                    this.Size = Settings.Default.Size;
                }
                else
                {
                    this.StartPosition = FormStartPosition.Manual;
                    this.DesktopBounds = new Rectangle(Settings.Default.WindowLocation, Settings.Default.Size);
                }
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveWindowLocation(); // Speichern der Fensterposition beim Schließen des Formulars
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            string appPath = "\"" + Application.ExecutablePath + "\""; // Fügen Sie Anführungszeichen hinzu

            RegistryKey rkApp = Registry.CurrentUser.OpenSubKey(RunKey, true);

            if (checkBox1.Checked)
            {
                // Wenn die Checkbox aktiviert ist, fügen Sie einen Eintrag in der Registrierung hinzu
                rkApp.SetValue(AppName, appPath);
            }
            else
            {
                // Wenn die Checkbox deaktiviert ist, entfernen Sie den Eintrag aus der Registrierung
                rkApp.DeleteValue(AppName, false);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            //SaveWindowLocation();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if(!showData)
            {
                showData = true;
            }
            else
            {
                showData = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            sendReset = true;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDownValue = numericUpDown1.Value;
            Settings.Default.NumericUpDownValue = numericUpDown1.Value;
            Settings.Default.Save();
        }

        private static List<object> collectHardwareInformation()
        {
            // OpenHardwareMonitor initialisieren
            Computer computer = new Computer
            {
                CPUEnabled = true,
                GPUEnabled = true,
                RAMEnabled = true,
                MainboardEnabled = true,
                FanControllerEnabled = true,
            };

            // Liste für die gesammelten Hardwareinformationen
            List<object> hardwareInfoList = new List<object>();

            computer.Open();
            computer.Accept(new UpdateVisitor());

            foreach (IHardware hardware in computer.Hardware)
            {
                foreach (ISensor sensor in hardware.Sensors)
                {
                    if (!serialPortopened)
                    {
                        Debug.WriteLine("Sensor Name: " + sensor.Name + " - Sensor Hw: " + sensor.Hardware + " - Sensor Type: " + sensor.SensorType + " - Sensor Value: " + sensor.Value);
                        Debug.WriteLine("Sensor Hw: " + sensor.Hardware);
                        Debug.WriteLine("Sensor Type: " + sensor.SensorType);
                        Debug.WriteLine("Sensor Value: " + sensor.Value);
                    }
                }

                foreach (IHardware subhardware in hardware.SubHardware)
                {
                    foreach (ISensor sensor in subhardware.Sensors)
                    {
                        if (sensor.SensorType == SensorType.Temperature && sensor.Name == "Temperature #1")
                            hardwareInfoList.Add(new { CT = customRound((float)sensor.Value) });

                        if (sensor.SensorType == SensorType.Fan && sensor.Name == "Fan #2")
                            hardwareInfoList.Add(new { CF = customRound((float)sensor.Value) });

                        if (sensor.SensorType == SensorType.Fan && sensor.Name == "Fan #3")
                            hardwareInfoList.Add(new { CaF = customRound((float)sensor.Value) });
                    }
                }

                foreach (ISensor sensor in hardware.Sensors)
                {
                    if (sensor.SensorType == SensorType.Load && sensor.Name.Contains("CPU Total"))
                        hardwareInfoList.Add(new { CU = customRound((float)sensor.Value) });

                    if (sensor.SensorType == SensorType.Load && sensor.Name == "GPU Core")
                        hardwareInfoList.Add(new { GU = customRound((float)sensor.Value) });

                    if (sensor.SensorType == SensorType.Fan && sensor.Name == "GPU")
                        hardwareInfoList.Add(new { GF = customRound((float)sensor.Value) });

                    if (sensor.SensorType == SensorType.Power && sensor.Name == "GPU Power")
                        hardwareInfoList.Add(new { GP = customRound((float)sensor.Value) });

                    if (sensor.SensorType == SensorType.Temperature && sensor.Name == "GPU Core")
                        hardwareInfoList.Add(new { GT = customRound((float)sensor.Value) });

                    if (sensor.SensorType == SensorType.Data && sensor.Name == "Used Memory")
                        hardwareInfoList.Add(new { RU = customRound((float)sensor.Value) });
                }
            }

            hardwareInfoList.Add(new { BL = NumericUpDownValue });
            if (sendReset)
            {
                hardwareInfoList.Add(new { RX = 1 });
                sendReset = false;
            }

            computer.Close();

            return hardwareInfoList;
        }

        static int customRound(float value)
        {
            // Überprüfen, ob der Dezimalteil größer oder gleich 0.5 ist
            // und entsprechend aufrunden oder abrunden
            if (value - Math.Floor(value) >= 0.5f)
                return (int)Math.Ceiling(value);
            else
                return (int)Math.Floor(value);
        }

        private void StartDataTransmission(SerialPort serialPort)
        {
            bool test = true;
            Thread thread = new Thread(() =>
            {
                while (true)
                {
                    string jsonData = JsonSerializer.Serialize(collectHardwareInformation());
                    
                    try
                    {
                        if(!serialPortopened) {
                            serialPort.Open();
                            serialPortopened = true;                            
                        }
                        if (test)
                        {
                            serialPort.Write(jsonData);
                            test = false;
                            AppendTextToTextBox(jsonData);
                        }
                        if (showData)
                        {
                            //AppendTextToTextBox(jsonData);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error while sending data: " + ex.Message);
                        Console.WriteLine("Reinitializing the COM connection.");
                        AppendTextToTextBox("Error while sending data: " + ex.Message);
                        AppendTextToTextBox("Reinitializing the COM connection.");
                        serialPortopened = false;
                        serialPort = connectToPort(this);
                    }
                    finally
                    {
                        Thread.Sleep(1000);
                        //serialPort.Close();
                    }
                    Thread.Sleep(100);
                }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        static SerialPort connectToPort(Form1 form)
        {
            try
            {
                int baudRate = 115200; // Beispiel: 9600 Baud
                SerialPort serialPort = new SerialPort(comPortInitialisize(form), baudRate);
                return serialPort;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        static string comPortInitialisize(Form1 form)
        {
            // Serielle Porteinstellungen
            string[] ports = SerialPort.GetPortNames();
            string comPortName = "COM1";
            // Durch jeden verfügbaren COM-Port iterieren und nach dem Controller suchen
            foreach (string port in ports)
            {
                if (IsControllerConnected(port, form))
                {
                    Console.WriteLine($"Controller is connected to {port}.");
                    form.AppendTextToTextBox($"Controller is connected to {port}.");
                    // Controller gefunden
                    comPortName = port;
                    break; // Brechen Sie die Schleife ab, wenn der Port gefunden wurde
                }
            }
            return comPortName;
        }

        // Funktion zum Überprüfen der Controller-Verbindung
        static bool IsControllerConnected(string portName, Form1 form)
        {
            try
            {
                using (SerialPort serialPort = new SerialPort(portName))
                {
                    // Versuchen, den seriellen Port zu öffnen, um zu sehen, ob der Controller angeschlossen ist
                    serialPort.Open();
                    Console.WriteLine($"The port {portName} has been opened.");
                    form.AppendTextToTextBox($"The port {portName} has been opened.");
                    return true; // Der Controller ist angeschlossen
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Der Port ist nicht verfügbar oder wird bereits von einem anderen Programm verwendet
                Console.WriteLine($"The port {portName} is not available or is already in use by another program!");
                form.AppendTextToTextBox($"The port {portName} is not available or is already in use by another program!");
                return false;
            }
            catch (Exception ex)
            {
                // Ein anderer Fehler ist aufgetreten
                Console.WriteLine($"Error opening port {portName}: {ex.Message}");
                form.AppendTextToTextBox($"Error opening port {portName}: {ex.Message}");
                return false;
            }
        }

        // Methode zum Hinzufügen von Text zur TextBox
        private void AppendTextToTextBox(string text)
        {
            if (textBox1.InvokeRequired)
            {
                textBox1.Invoke((MethodInvoker)(() => AppendTextToTextBox(text)));
            }
            else
            {
                textBox1.AppendText(text + Environment.NewLine);
            }
        }
    }
}

public class UpdateVisitor : IVisitor
{
    public void VisitComputer(IComputer computer)
    {
        computer.Traverse(this);
    }
    public void VisitHardware(IHardware hardware)
    {
        hardware.Update();
        foreach (IHardware subHardware in hardware.SubHardware) subHardware.Accept(this);
    }
    public void VisitSensor(ISensor sensor) { }
    public void VisitParameter(IParameter parameter) { }
}

public class HardwareInfo
{
    public int Index { get; set; }
    public float Value { get; set; }
}
