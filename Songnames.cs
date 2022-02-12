using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Text.Json;
using System.IO;

namespace PPTXcreator
{
    public static class Songnames
    {
        public static Dictionary<string, string> Names;

        /// <summary>
        /// Load the song names from the embedded file to <see cref="Names"/>
        /// </summary>
        public static void LoadNames()
        {
            // Get a stream of the embedded file
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream filestream = assembly.GetManifestResourceStream("PPTXcreator.Resources.liednamen"))
            {
                // Get the dictionary from the JSON formatted file and save it to Songnames.Names
                JsonDocument songnames = JsonDocument.Parse(filestream);
                Names = songnames.RootElement.Deserialize<Dictionary<string, string>>();
            }
        }

        /// <summary>
        /// Return the value of <see cref="Names"/>[<paramref name="songnumber"/>], or null if not found
        /// </summary>
        public static string GetName(string songnumber)
        {
            if (Names is null) return null;
            try
            {
                return Names[songnumber];
            }
            catch (KeyNotFoundException)
            {
                return null;
            }
        }
    }
}
