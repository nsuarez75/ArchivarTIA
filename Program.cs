using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Siemens.Engineering;
using System.IO;
using Microsoft.Win32;
using Siemens.Engineering.HW;
using System.Windows;
using System.Globalization;
using System.Threading;

namespace ArchivarTIA
{
    class Program
    {
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
                TiaPortal tia = new TiaPortal(TiaPortalMode.WithoutUserInterface);
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

        static void Main(string[] args)
        {

            Console.SetWindowSize(60, 10);
            List<string> proyectos = BuscarProyectos();
            ArchivarProyectos(proyectos);

            
        }
    }
}
