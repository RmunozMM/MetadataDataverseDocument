using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using MetadataDataverseDocument.UI;

namespace MetadataDataverseDocument
{
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "Metadata Dataverse Document"),
        ExportMetadata("Description", "Document entity relationships, metadata, schema, alternate keys and ERD diagrams from Dataverse solutions into Excel & Mermaid"),
        ExportMetadata("PluginType", "Documentation"),
        ExportMetadata("SmallImageBase64", "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"),
        // ATENCION: el base64 de esta imagen debe medir MENOS de 16383 caracteres.
        // El formato de longitud comprimida de los blobs de atributos (ECMA-335) usa 2 bytes
        // hasta 0x3FFF = 16383; pasado ese punto exige el formato de 4 bytes, y el compilador
        // de Mono (mcs) lo codifica mal: escribe 2 bytes con un valor truncado. El blob queda
        // inconsistente y .NET Framework lanza CustomAttributeFormatException al leerlo, lo que
        // ocurre durante la composicion MEF de XrmToolBox al arrancar y TUMBA LA CARGA DE TODOS
        // LOS PLUGINS, no solo de este. La imagen actual mide 14236 caracteres.
        ExportMetadata("BigImageBase64", "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"),
        ExportMetadata("BackgroundColor", "Lavender"),
        ExportMetadata("PrimaryFontColor", "Black"),
        ExportMetadata("SecondaryFontColor", "Gray")]
    public class Plugin : PluginBase
    {
        // AppDomain.CurrentDomain.AssemblyResolve es un evento de TODO EL PROCESO, no del
        // plugin. XrmToolBox puede instanciar esta clase mas de una vez, y antes cada
        // instancia agregaba OTRO handler que nunca se quitaba: el resultado eran varios
        // handlers nuestros interceptando las resoluciones de assembly de todos los demas
        // plugins cargados. Se registra una sola vez por proceso.
        private static int _resolverRegistered;

        public Plugin()
        {
            if (Interlocked.Exchange(ref _resolverRegistered, 1) == 0)
            {
                AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolveEventHandler;
            }
        }

        public override IXrmToolBoxPluginControl GetControl()
        {
            return new MetadataDocumentControl();
        }

        /// <summary>
        /// Resuelve UNICAMENTE las dependencias de este plugin, y solo desde su propia
        /// subcarpeta.
        ///
        /// Por que importa tanto: este handler se llama para cada assembly que el runtime no
        /// logra resolver en TODO el proceso de XrmToolBox, incluidas las peticiones de otros
        /// plugins. Si respondemos a una peticion ajena con NUESTRA copia de una libreria
        /// compartida (EPPlus, Microsoft.Xrm.Sdk, System.Resources.Extensions...), ese otro
        /// plugin recibe una version distinta de la que espera y falla con
        /// FileLoadException / TypeLoadException / MissingMethodException.
        ///
        /// La version anterior de este metodo componia la ruta con "$argName.dll" -
        /// interpolacion de PowerShell dentro de un literal de C#-, asi que buscaba un archivo
        /// llamado literalmente "$argName.dll" y nunca encontraba nada. Eso obligaba a copiar
        /// las dependencias a la RAIZ de la carpeta Plugins para que el runtime las hallara por
        /// sondeo normal, y es justamente esa copia en la raiz la que queda visible para todos
        /// los demas plugins. Al arreglar la ruta, las dependencias vuelven a la subcarpeta y
        /// dejan de interferir.
        /// </summary>
        private static Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            try
            {
                var thisAssembly = typeof(Plugin).Assembly;

                // Si sabemos quien pide y no somos nosotros, no es asunto nuestro.
                if (args.RequestingAssembly != null && args.RequestingAssembly != thisAssembly)
                {
                    return null;
                }

                var requested = new AssemblyName(args.Name);

                // Solo dependencias que este assembly realmente declara.
                bool isOwnDependency = thisAssembly
                    .GetReferencedAssemblies()
                    .Any(a => string.Equals(a.Name, requested.Name, StringComparison.OrdinalIgnoreCase));
                if (!isOwnDependency)
                {
                    return null;
                }

                // Nuestras dependencias viven en Plugins\<NombreDelAssembly>\ y en ningun otro
                // lado. Deliberadamente NO se busca en la raiz de Plugins: lo que hay ahi es
                // territorio compartido con el resto de los plugins.
                string pluginsDir = Path.GetDirectoryName(thisAssembly.Location);
                string ownFolder = Path.GetFileNameWithoutExtension(thisAssembly.Location);
                string candidate = Path.Combine(pluginsDir, ownFolder, requested.Name + ".dll");

                if (!File.Exists(candidate))
                {
                    return null;
                }

                // Nunca devolver una version mas antigua que la solicitada: si alguien pide una
                // version mayor de una libreria compartida, dejamos que el runtime siga buscando
                // en vez de entregarle la nuestra y romperlo.
                if (requested.Version != null)
                {
                    var candidateName = AssemblyName.GetAssemblyName(candidate);
                    if (candidateName.Version != null && candidateName.Version < requested.Version)
                    {
                        return null;
                    }
                }

                return Assembly.LoadFrom(candidate);
            }
            catch
            {
                // Una excepcion lanzada desde un handler de AssemblyResolve se propaga a quien
                // haya disparado la carga, que puede ser un plugin ajeno. Fallar en silencio y
                // dejar que el runtime continue con su propio sondeo.
                return null;
            }
        }
    }
}
