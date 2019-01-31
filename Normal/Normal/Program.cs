using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Permissions;
using Microsoft.Win32;



namespace EnableExcelAddin
{
    class MIORegKeyManager
    {
        static void Main(string[] args)
        {
            // Path that locates the addins
            string registrypath = "Software\\Microsoft\\Office\\14.0\\Excel\\Options";

            MIORegKeyManager obj = new MIORegKeyManager();
            RegistryKey currentuser64 = null;
            RegistryKey rootfile = null;
            try
            {// Opens a new RegistryKey that represents the path on the local machine with the specified view
                currentuser64 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);
                //Console.WriteLine(currentuser64);
            }

            catch (ArgumentException e)
            {
                Console.WriteLine("The hKey or view is invalid. Stacktrace:"+e.Message);
                 Environment.Exit(1);
               // throw new Exception("The hKey or view is invalid. Stacktrace:"+e.Message);
            }


            try
            {
                // Retrieves the subkey as read-only
                rootfile = currentuser64.OpenSubKey(registrypath, true);
                //Console.WriteLine(rootfile);
            }

            catch (ObjectDisposedException e)
            {
                Console.WriteLine("The RegistryKey is closed.");
                Environment.Exit(1);
            }

            catch (ArgumentNullException e)
            {
                Console.WriteLine("The requested subkey is null.");
                Environment.Exit(1);
            }

            // Sets the values of the disabled addins

            obj.SetValues(rootfile, obj.GetNames(rootfile));
            Console.WriteLine(rootfile);
            //Console.ReadLine();
            Environment.Exit(0);
        }


        public string[] GetNames(RegistryKey rootfile)
        {
            string[] rfile =null;
            try
            {// Returns the names of the addins
                rfile = rootfile.GetValueNames(); 
            }

            catch (NullReferenceException nre)
            {
                Console.WriteLine("The path to the addins or the file with the addins is null.");
                Environment.Exit(1);
            }
            return rfile;
        }

        public void SetValues(RegistryKey rootfile, string[] nameList)
        {
            foreach (string valueName in nameList)
            {
                String value = null;
                try
                { //Retrieves the value from the addin
                    value = rootfile.GetValue(valueName).ToString();
                }

                catch (NullReferenceException nre)
                {
                    Console.WriteLine("The value is null.");
                    Environment.Exit(1);
                }

                catch (System.IO.IOException e)
                {
                    Console.WriteLine("The RegistryKey that contains the specified value has been marked for deletion.");
                    Environment.Exit(1);
                }

                // Prints the information from the HKEY_CURRENT_USER subkey
                Console.WriteLine("{0,-8}: {1}", valueName, value);

                // When valuename is like OPEN and value is empty 
                if (valueName.Contains("OPEN") && string.IsNullOrEmpty(value))
                {
                    // Value of addin if it is disabled
                    value = @"""C:\\Users\\harshvir randhawa\\AppData\\Local\\MIOExcelAddIn\\MIOExcelAddIn\\app-1.0.1.12\\MIO.Excel.AddIn.xll""";

                    // Prints the new information from the HKEY_CURRENT_USER subkey
                    Console.WriteLine(value);

                    try
                    {//Sets value of addin if it is disabled
                        rootfile.SetValue(valueName, value);
                    }

                    catch (System.IO.IOException e)
                    {
                        Console.WriteLine("The RegistryKey object represents a root-level node, and the operating system is Windows 2000, Windows XP, or Windows Server 2003.");
                        Environment.Exit(1);
                    }

                }

            }
        }
    
    }
}