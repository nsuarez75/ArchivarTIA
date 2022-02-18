using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Siemens.Engineering;
using System.Threading;

namespace ArchivarTIA
{
    class Program
    {
        private static TiaPortal tia;
        private static string plc;
        private static string hmi;

        

        static void Main(string[] args)
        {
            Console.SetWindowSize(60, 10);
            Console.Title = "ArchivarTIA";
            //Subscribe to Assembly resolve event. Event fires when any assembly binding fails.
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;

            //Path to registry is HKEY_LOCAL_MACHINE\SOFTWARE\Siemens\Automation\Openness\*version\PublicAPI\*assembly
            List<string> versions = RegistryReader.GetVersions();

            //I always have only one version installed on my VMs so i can keep this fixed, you could read the versions list to retrieve other installed versions and choose which one you want
            int nr = 1;
            List<string> assemblies = RegistryReader.GetAssemblies(versions[nr - 1]);
            RegistryReader.GetAssemblyPath(versions[nr - 1], assemblies.Last(), out plc, out hmi);

            List<string> projects = FindProjects();
            if (projects != null) ArchiveProjectOnAssemblyResolve(projects);

        }

        /// <summary>
        /// Returns Project list found in Desktop
        /// </summary>
        /// <returns></returns>
        static List<string> FindProjects()
        {
            string targetDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            List<string> directories = new List<string>(Directory.EnumerateDirectories(targetDirectory));
            List<string> projects = new List<string>();
            Int16 count = 0;
            List<string> versions = new List<string>() { ".ap13",".ap13_1",".ap14",".ap15",".ap15_1",".ap16",".ap17"};
            Console.WriteLine($"Searching projects in {targetDirectory}");
            foreach (var dir in directories)
            {
                List<string> files = new List<string>(Directory.EnumerateFiles(dir));
                foreach (var file in files)
                {
                    if (versions.Contains(Path.GetExtension(file)))
                    {
                        projects.Add(Path.GetFullPath(file));
                        count++;
                    }
                   
                }
            }

            if (count > 1) Console.WriteLine($"{count} projects found");
            else if (count == 1) Console.WriteLine($"{count} project found");
            else Console.WriteLine("Couldn't find any project");

            return projects;
        }

        /// <summary>
        /// Archive the projects located in the project list <paramref name="projects"/> recibida 
        /// </summary>
        /// <param name="projects"></param>
        static void ArchiveProjectOnAssemblyResolve(List<string> projects)
        {
            try
            {
                Console.WriteLine("Loading hidden TIA instance");
                tia = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                ProjectComposition tiaProjects = tia.Projects;

                foreach (var project in projects)
                {
                    Console.Clear();
                    string projectName = Path.GetFileNameWithoutExtension(project);
                    DateTime moment = DateTime.Now;
                    string date = $"{moment.Year}{moment.Month}{moment.Day}_{moment.Hour}{moment.Minute}";
                    projectName = projectName + "_" + date;
                    string projectExtension = Path.GetExtension(project);
                    projectName = projectName + projectExtension.Insert(1,"z");

                    try
                    {
                        FileInfo directorio = new FileInfo(project);
                        Console.WriteLine($"Archiving project {Path.GetFileNameWithoutExtension(project)}");
                        Project tiaProject = tiaProjects.Open(directorio);
                        DirectoryInfo target = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                        tiaProject.Archive(target, $"{projectName}", ProjectArchivationMode.Compressed);
                        tiaProject.Close();
                        Console.WriteLine($"{projectName} archived!");
                        Thread.Sleep(1000);
                    }

                    catch (Exception e)
                    {
                        Console.WriteLine($"Error archiving project {projectName}");
                        Console.WriteLine(e);
                        Console.ReadKey();
                    }
                }

            }

            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadKey();
            }

        }

        public static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            var assemblyName = new AssemblyName(args.Name);
            string path = "";
            //Check for referenced Siemens assemblies.
            if (assemblyName.Name.EndsWith("Siemens.Engineering"))
                path = plc;
            if (assemblyName.Name.EndsWith("Siemens.Engineering.Hmi"))
                path = hmi;
            //Load Selected Siemens Assembly.
            if (string.IsNullOrEmpty(path) == false
                && File.Exists(path))
            {
                return Assembly.LoadFrom(path);
            }
            // return null for all other assemblies.
            return null;
        }
    }
}
