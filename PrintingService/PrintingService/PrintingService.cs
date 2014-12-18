using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;

using System.IO;
using System.Net.Sockets;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Printing;
using System.Drawing.Printing;
using System.Text.RegularExpressions;

namespace PrintingService
{

    public partial class PrintingService : ServiceBase
    {
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);
                
        private string disable = "", current_server = "", location_property_name = "", working_directory = "", download_path = "", key_value_filename = "", printer_list_filename = "";
        private string[] servers = new string[] { };

        /**
         * Read in the configuration file containing runtime parameters.
         * */
        void ParseConfigFile()
        {
            //string disable = "", current_server = "", location_property_name = "", printer_list_path = "";
            //string[] servers = new string[]{};

            // Get print servers from config file
            if (File.Exists(@"C:\Tools\Printing\config.txt"))
            {
                System.IO.StreamReader file = new System.IO.StreamReader(@"C:\Tools\Printing\config.txt");
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    string[] words;
                    if (line.Contains("[server_list]"))
                    {
                        string[] delim = new string[] { "= " };
                        words = line.Split(delim, StringSplitOptions.RemoveEmptyEntries);
                        servers = words[1].Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                        //Console.WriteLine("PrintingService.PrimaryServer: " + primary_server);
                    }
                    if (line.Contains("[disable_service]"))
                    {
                        words = line.Split((string[])null, 3, StringSplitOptions.RemoveEmptyEntries);
                        disable = words[2];
                        if (disable.CompareTo("1") == 0)
                        {
                            Console.WriteLine("PrintingService Disabling service... " + disable);
                            return;
                        }
                    }
                    if (line.Contains("[location_property_name]"))
                    {
                        words = line.Split((string[])null, 3, StringSplitOptions.RemoveEmptyEntries);
                        location_property_name = words[2];
                    }
                    if (line.Contains("[working_directory]"))
                    {
                        words = line.Split((string[])null, 3, StringSplitOptions.RemoveEmptyEntries);
                        working_directory = words[2];
                    }
                    if (line.Contains("[download_path]"))
                    {
                        words = line.Split((string[])null, 3, StringSplitOptions.RemoveEmptyEntries);
                        download_path = words[2];
                    }
                    if (line.Contains("[key_value_filename]"))
                    {
                        words = line.Split((string[])null, 3, StringSplitOptions.RemoveEmptyEntries);
                        key_value_filename = words[2];
                    }
                    if (line.Contains("[printer_list_filename]"))
                    {
                        words = line.Split((string[])null, 3, StringSplitOptions.RemoveEmptyEntries);
                        printer_list_filename = words[2];
                    }
                }
                file.Close();
            }
        }

        /**
         * Setup a printer for every printer listed in the printers list file.  Checks for updates to the printer configurations and if necessary downloads the updates
         * 
         * */
        void PrinterSetup()
        {
            // Parse the config file
            ParseConfigFile();

            // Get a list of all printers that should be connected to this machine.  For each printer, download the key value file from the first server available, and 
            // also has the print service available for that printer on that particular server.  If the service is down then download the key value file from the next
            // server in the list.  This maintains each print server as an island unto itself as far as configuration for any particular printer is concerned.
            if (File.Exists(working_directory + printer_list_filename))
            {
                // Holds a list of our printer names that are installed locally on this machine
                List<string> installedPrinters = GetInstalledNetworkPrinters(location_property_name);

                // Delete any of our printers that are not within the printers.txt list
                DeletePrinters(installedPrinters);

                // Default printer 
                string default_printer = "";

                System.IO.StreamReader file = new System.IO.StreamReader(working_directory + printer_list_filename);
                string line;
                int count = 0;
                while ((line = file.ReadLine()) != null)
                {
                    if (count == 0)
                        default_printer = line;
                    foreach (string server in servers)
                    {
                        // Download the key value printer => port file from the Server
                        WebClient client = new WebClient();
                        string url = "http://" + server + "/printing/" + key_value_filename;
                        try
                        {
                            client.DownloadFile(new Uri(url), download_path + key_value_filename);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message + ": " + server);
                        }

                        // Get the port associated with the current printer key
                        if (File.Exists(download_path + key_value_filename))
                        {
                            System.IO.StreamReader keyFile = new System.IO.StreamReader(download_path + key_value_filename);
                            string key, value;
                            while ((key = keyFile.ReadLine()) != null)
                            {
                                // Search for value
                                if (key.Contains(line))
                                {
                                    // Split at blank spaces
                                    string[] split = key.Split((string[])null, 2, StringSplitOptions.RemoveEmptyEntries);
                                    value = split[1];

                                    // Check that the printing service is available on the current server for the given port and printer
                                    if (PrintingServiceCheck(server, Convert.ToInt32(value)))
                                    {
                                        // Found the printing service requested by port number
                                        current_server = server;
                                        break;
                                    }
                                }
                            }
                            keyFile.Close();
                        }
                    }
                    if (current_server.CompareTo("") == 0)
                        Console.WriteLine("No printing service available for printer: " + line);
                    else
                    {
                        WebClient client = new WebClient();
                        string url = "";

                        // Printing service found on one of the servers, download the MD5 file and check if we need to download a new copy of the BRM file
                        Console.WriteLine("Printing service is available for: " + line);
                        Console.WriteLine("Checking for an updated printer configuration...");

                        // If either the printers BRM or MD5 file do not exist locally then download both files from the server
                        if (!File.Exists(working_directory + line + ".brm") || !File.Exists(working_directory + line + ".md5"))
                        {
                            try
                            {
                                url = "http://" + current_server + "/printing/" + line + ".md5";
                                client.DownloadFile(new Uri(url), working_directory + line + ".md5");

                                url = "http://" + current_server + "/printing/" + line + ".brm";
                                client.DownloadFile(new Uri(url), working_directory + line + ".brm");
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Unable to download file " + line + " from server " + current_server + ": " + e.Message);
                            }
                        }
                        else
                        {
                            url = "http://" + current_server + "/printing/" + line + ".md5";
                            bool updateNeeded = true;
                            try
                            {
                                client.DownloadFile(new Uri(url), download_path + line + ".md5");

                                if (File.Exists(download_path + line + ".md5"))
                                {
                                    System.IO.StreamReader md5File = new System.IO.StreamReader(download_path + line + ".md5");
                                    string md5Line;
                                    while ((md5Line = md5File.ReadLine()) != null)
                                    {
                                        md5Line = md5Line.Split(new string[] { "  " }, StringSplitOptions.RemoveEmptyEntries)[0];
                                        using (var md5 = MD5.Create())
                                        {
                                            string hash = BitConverter.ToString(md5.ComputeHash(File.ReadAllBytes(working_directory + line + ".brm"))).Replace("-", "");

                                            if (md5Line.ToUpper().CompareTo(hash.ToUpper()) == 0)
                                            {
                                                Console.WriteLine("Configuration matched!  Proceeding with installation . . .");
                                                updateNeeded = false;
                                            }
                                            else
                                            {
                                                // They do not match so download the updated versions of the BRM and MD5 files for this printer
                                                Console.WriteLine("Configuration match failed!  Downloading updated configuration . . .");
                                            }
                                        }
                                    }
                                    md5File.Close();
                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            if (updateNeeded)
                            {
                                // Download MD5 file
                                url = "http://" + current_server + "/printing/" + line + ".md5";
                                client.DownloadFile(new Uri(url), working_directory + line + ".md5");

                                // Download BRM file
                                url = "http://" + current_server + "/printing/" + line + ".brm";
                                client.DownloadFile(new Uri(url), working_directory + line + ".brm");

                                // Install printer
                                InstallPrinter(line);
                            }
                            else
                            {
                                // File versions passed checks so if the printer already exists on the list of installed printers do nothing, otherwise install it
                                if (!installedPrinters.Contains(line))
                                {
                                    // Install printer
                                    InstallPrinter(line);
                                }
                            }
                        }
                    }
                    count++;
                }
                file.Close();

                // Set default printer
                SetDefaultPrinter(default_printer);
            }
            else
            {
                Console.WriteLine("Printers.txt file not found!");
            }
        }

        /**
         * Retrieve a list of currently installed network printers that have been previously installed through the Marriot Library system.  These printers
         * will be identified by a comment property string as listed in the printers configuration file located in the Printers folder.  This list can then be 
         * used to uninstall printers that are no longer needed and have been removed from the printers.txt list file in the Printers folder.
         * 
         * @param   string          The printer comment property string used to identify printers previously installed on the system through this Service
         * @ret     List<string>    A List of printer names 
         * */
        List<string> GetInstalledNetworkPrinters(string location_property_name)
        {
            // The list of printers to return
            List<string> printers = new List<string>();

            // Search all installed printers and retrieve the comment property string that will identify which printers this service previously installed on the system
            System.Management.ManagementObjectSearcher searcher = new System.Management.ManagementObjectSearcher("SELECT * FROM Win32_Printer");
            foreach (System.Management.ManagementObject printer in searcher.Get())
            {
                string location = "";
                if (printer["Location"] != null)
                    location = printer["Location"].ToString();
                if (location.Contains(location_property_name) && printer["Location"] != null)
                {
                    // This is one of our network printers
                    printers.Add(printer["Name"].ToString());
                }
            }

            // Return the list
            return printers;
        }

        /**
         * Look for any currently installed printers that were installed through this Service, but are no longer in the printers list file.  Those printers should
         * be removed from the system.
         * 
         * @param   List<string>    A list of currently installed printer names that were installed through this Service
         * */
        void DeletePrinters(List<string> installedPrinters)
        {
            // For each printer in the installed List<> if it is NOT in the printers list, delete it
            List<string> lPrinters = new List<string>();
            if (File.Exists(working_directory))
            {
                System.IO.StreamReader file = new System.IO.StreamReader(working_directory);
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    lPrinters.Add(line);
                }
                file.Close();
            }
            foreach (string iPrinter in installedPrinters)
            {
                if (!lPrinters.Contains(iPrinter))
                {
                    // Delete
                    System.Management.ManagementScope oManagementScope = new System.Management.ManagementScope(System.Management.ManagementPath.DefaultPath);
                    oManagementScope.Connect();
                    System.Management.SelectQuery query = new System.Management.SelectQuery("SELECT * FROM Win32_Printer");
                    System.Management.ManagementObjectSearcher search = new System.Management.ManagementObjectSearcher(oManagementScope, query);
                    System.Management.ManagementObjectCollection printers = search.Get();
                    foreach (System.Management.ManagementObject printer in printers)
                    {
                        string pName = printer["Name"].ToString().ToLower();
                        if (pName.Equals(iPrinter.ToLower()))
                        {
                            printer.Delete();
                            break;
                        }
                    }
                }
            }

        }

        /**
         * Install a given printer using native Windows BRM utility.
         * 
         * @param   string  The name of the printer being installed
         * */
        void InstallPrinter(string printer)
        {
            string command = @"C:\Windows\System32\spool\tools\PrintBrm.exe";
            ProcessStartInfo startInfo = new ProcessStartInfo(command);
            startInfo.WindowStyle = ProcessWindowStyle.Minimized;
            Process.Start(startInfo);
            startInfo.Arguments = @"-r -f C:\Tools\Printing\" + printer + ".brm -noacl -o force";
            Process.Start(startInfo);
        }

        /**
         * Determines if the printing service is available on the given server and port.  
         * 
         * @param   string  Name of the server the service should be running on
         * @param   int     The port the printing service should be listening on
         * 
         * @ret     bool    Success or failure depending on if the port is listening on the server
         * */
        bool PrintingServiceCheck(string server_name, int port)
        {
            Console.WriteLine("Checking printing service status on " + server_name + " . . .");
            TcpClient client = new TcpClient();
            try
            {
                if (server_name != "")
                    client.Connect(server_name, port);
                return true;
            }
            catch (Exception)
            {
                // printing server service check failed
                Console.WriteLine("Printing service not available!");
                return false;
            }
        }

        public PrintingService()
        {
            // Set this property to true so we can detect logon events
            CanHandleSessionChangeEvent = true;            

            InitializeComponent();
        }

        /**
         * Listen for the SessionChange Event and detect system logons.  When a logon is detected begin running the printer setup.
         * */
        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            base.OnSessionChange(changeDescription);
            
            switch (changeDescription.Reason)
            {
                case SessionChangeReason.SessionLogon:
                    //EventLog.WriteEntry("PrintingService.OnSessionChange: Logon");

                    // Setup printers
                    PrinterSetup();
                    break;
            }
        }

        protected override void OnStart(string[] args)
        {
        }

        protected override void OnStop()
        {
        }
    }
}
