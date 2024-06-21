using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;

namespace Setup
{
    internal class Program
    {
        static void Main(string[] args)
        {


            try
            {

                // Create new command and set its properites and then Register it
                    Command myCommand = new Command();
                    myCommand.projectName = "CImatronAPI_Template";
                    myCommand.className = "Command";
                    myCommand.IconName = "Icon.ico";
                    myCommand.dllPath = @"CImatronAPI_Template.dll";
                    Command.registerCommand(myCommand);
 
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }

            Console.WriteLine("Press Enter to exit.");
            Console.ReadLine();


        }
        public class Command
        {
            //Command Properties

            //name of the icon in the directory example("Icon.ico")
            public string @IconName {get; set; }

            //The name of the Project or Namespace
            public string projectName { get; set; }

            // The name of the class where the command belongs in this example it is just named command
            public string className { get; set; }

            // The path to the dll From the app directroy

            public string dllPath { get; set; }


            // Defaults the Cimatron Program Folder but you can set it where it needs to be manually
            public static void registerCommand(Command command, string cimatronProgramFolder = @"C:\Program Files\Cimatron\Cimatron") {


            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string IconPath = Path.Combine(currentDirectory, command.IconName);
            string fullDllPath = System.IO.Path.Combine(currentDirectory, command.dllPath);

                string regasmPath = @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe";






                ProcessStartInfo unregister = new ProcessStartInfo
                {
                    FileName = regasmPath,
                    Arguments = $@" /u ""{fullDllPath}""",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(unregister))
                {
                    process.WaitForExit();
                    Console.WriteLine(process.StandardOutput.ReadToEnd());
                }

                ProcessStartInfo RegisterDll = new ProcessStartInfo
                {
                    FileName = regasmPath,
                    Arguments = $@"""{fullDllPath}"" /codebase /tlb",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (Process process = Process.Start(RegisterDll))
                {
                    process.WaitForExit();
                    Console.WriteLine(process.StandardOutput.ReadToEnd());
                }

                string RcommandArguments = "\"" + @fullDllPath + "\" " + command.projectName + " " + command.className + " " + "\"" + @IconPath + "\"";
                
                if (Directory.Exists(cimatronProgramFolder))
                {
                    // Get an array of all subfolder paths within the specified folder
                    string[] CimVersions = Directory.GetDirectories(cimatronProgramFolder);

                    // Iterate through each subfolder path and print it
                    if (CimVersions.Length > 0)
                    {
                        foreach (string CimVersion in CimVersions)
                        {
                            try
                            {
                                Console.WriteLine(CimVersion);
                                string registerApiCommandsExePath = System.IO.Path.Combine(CimVersion, "Program", "Register_API_Commands.exe");


                                ProcessStartInfo registerApiCommand = new ProcessStartInfo
                                {
                                    FileName = registerApiCommandsExePath,
                                    Arguments = RcommandArguments,
                                    Verb = "runas", // Request admin privileges
                                    UseShellExecute = true
                                };
                                Process.Start(registerApiCommand)?.WaitForExit(); // Register the DLL
                                Console.WriteLine("API commands registered successfully.");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to register command class {command.className}\n {ex.Message}");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("The specified folder does not exist.");
                }
               

            }

        }
    }
}
