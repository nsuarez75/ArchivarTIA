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

            int nr = 1;
            List<string> assemblies = RegistryReader.GetAssemblies(versions[nr - 1]);
            RegistryReader.GetAssemblyPath(versions[nr - 1], assemblies.Last(), out plc, out hmi);

            List<string> proyectos = BuscarProyectos();
            ArchivarProyectos1(proyectos);
        }

        /// <summary>
        /// Devuelve una lista de paths a los proyectos encontrados
        /// </summary>
        /// <returns></returns>
        static List<string> BuscarProyectos()
        {
            string targetDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            List<string> directorios = new List<string>(Directory.EnumerateDirectories(targetDirectory));
            List<string> proyectos = new List<string>();
            Int16 count = 0;

            Console.WriteLine("Buscando proyectos");
            foreach (var dir in directorios)
            {
                List<string> archivos = new List<string>(Directory.EnumerateFiles(dir));
                foreach (var archivo in archivos)
                {
                    if (Path.GetExtension(archivo) == ".ap16")
                    {
                        proyectos.Add(Path.GetFullPath(archivo));
                        count++;
                    }
                    else if (Path.GetExtension(archivo) == ".ap15")
                    {
                        proyectos.Add(Path.GetFullPath(archivo));
                        count++;
                    }
                    else if (Path.GetExtension(archivo) == ".ap15_1")
                    {
                        proyectos.Add(Path.GetFullPath(archivo));
                        count++;
                    }
                }
            }

            if (count > 1) Console.WriteLine($"Se han encontrado {count} proyectos");
            else Console.WriteLine($"Se ha encontrado {count} proyecto");

            return proyectos;
        }

        /// <summary>
        /// Archiva los proyectos que se encuentran en la lista de paths <paramref name="proyectos"/> recibida 
        /// </summary>
        /// <param name="proyectos"></param>
        static void ArchivarProyectos(List<string> proyectos)
        {
            try
            {
                Console.WriteLine("Abriendo TIA Portal");
                tia = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                ProjectComposition tiaProjects = tia.Projects;

                foreach (var proyecto in proyectos)
                {
                    Console.Clear();
                    string nombreProyecto = Path.GetFileNameWithoutExtension(proyecto);
                    DateTime moment = DateTime.Now;
                    string fecha = $"{moment.Year}{moment.Month}{moment.Day}_{moment.Hour}{moment.Minute}";
                    nombreProyecto = nombreProyecto + "_" + fecha;
                    if (Path.GetExtension(proyecto) == ".ap16") nombreProyecto = nombreProyecto + ".zap16";
                    else if (Path.GetExtension(proyecto) == ".ap15") nombreProyecto = nombreProyecto + ".zap15";
                    else if (Path.GetExtension(proyecto) == ".ap15_1") nombreProyecto = nombreProyecto + ".zap15_1";


                    try
                    {
                        FileInfo directorio = new FileInfo(proyecto);
                        Console.WriteLine($"Archivando proyecto {Path.GetFileNameWithoutExtension(proyecto)}");
                        Project project = tiaProjects.Open(directorio);
                        DirectoryInfo target = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                        project.Archive(target, $"{nombreProyecto}", ProjectArchivationMode.Compressed);
                        project.Close();
                        Console.WriteLine($"{nombreProyecto} archivado correctamente");
                        Thread.Sleep(1000);
                    }

                    catch (Exception e)
                    {
                        Console.WriteLine($"Imposible archivar proyecto {nombreProyecto}");
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

        static void ArchivarProyectos1(List<string> proyectos)
        {
            try
            {
                Console.WriteLine("Abriendo TIA Portal");
                tia = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                ProjectComposition tiaProjects = tia.Projects;

                foreach (var proyecto in proyectos)
                {
                    Console.Clear();
                    string nombreProyecto = Path.GetFileNameWithoutExtension(proyecto);
                    DateTime moment = DateTime.Now;
                    string fecha = $"{moment.Year}{moment.Month}{moment.Day}_{moment.Hour}{moment.Minute}";
                    nombreProyecto = nombreProyecto + "_" + fecha;
                    if (Path.GetExtension(proyecto) == ".ap16") nombreProyecto = nombreProyecto + ".zap16";
                    else if (Path.GetExtension(proyecto) == ".ap15") nombreProyecto = nombreProyecto + ".zap15";
                    else if (Path.GetExtension(proyecto) == ".ap15_1") nombreProyecto = nombreProyecto + ".zap15_1";


                    try
                    {
                        FileInfo directorio = new FileInfo(proyecto);
                        Console.WriteLine($"Archivando proyecto {Path.GetFileNameWithoutExtension(proyecto)}");
                        Project project = tiaProjects.Open(directorio);
                        DirectoryInfo target = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                        project.Archive(target, $"{nombreProyecto}", ProjectArchivationMode.Compressed);
                        project.Close();
                        Console.WriteLine($"{nombreProyecto} archivado correctamente");
                        Thread.Sleep(1000);
                    }

                    catch (Exception e)
                    {
                        Console.WriteLine($"Imposible archivar proyecto {nombreProyecto}");
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
