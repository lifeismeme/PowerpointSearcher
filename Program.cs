using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace PowerpointSearcher
{
    class Program
    {
        static void Main(string[] args)
        {
            //string zipPath = @".\result.zip";

            //Console.WriteLine("Provide path where to extract the zip file:");
            //string extractPath = Console.ReadLine();

            //// Normalizes the path.
            //extractPath = Path.GetFullPath(extractPath);

            //// Ensures that the last character on the extraction path
            //// is the directory separator char.
            //// Without this, a malicious zip file could try to traverse outside of the expected
            //// extraction path.
            //if (!extractPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            //{
            //    extractPath += Path.DirectorySeparatorChar;
            //}

            string searchText;
            if(args.Length > 0)
            {
                searchText = args[0].ToLower();
            } else
            {
                searchText = "";
            }
            
            if(searchText != "")
            {
                string zipPath;
                while ((zipPath = Console.ReadLine()) != null)
                {
                    if (zipPath.EndsWith(".pptx"))
                    {
                        using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                        {
                            bool foundText = false;
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                if (foundText)
                                {
                                    break;
                                }

                                //filter for right files
                                if (!entry.FullName.StartsWith("ppt/slides/slide"))
                                {
                                    continue;
                                }
                                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                //using (StreamReader stream = new StreamReader(entry.Open()))
                                //{
                                //    string line;
                                //    while ((line = stream.ReadLine()) != null)
                                //    {
                                //        if (line.ToLower().Contains(searchText))
                                //        {
                                //            Console.WriteLine(zipPath);
                                //            foundText = true;
                                //            break;
                                //        }
                                //    }
                                //}

                                //XmlReaderSettings settings = new XmlReaderSettings();
                                //settings.Async = false;

                                using (XmlReader reader = XmlReader.Create(entry.Open()))
                                {
                                    reader.Read(); //skip the line metadata thing, `version="1.0".....
                                    while (reader.Read())
                                    {
                                        if (reader.Value == "" || reader.Value.Trim() == "")
                                        {
                                            continue;
                                        }
                                        if (reader.Value.ToLower().Contains(searchText))
                                        {
                                            Console.WriteLine(zipPath);
                                            foundText = true;
                                            break;
                                        }
                                    }
                                }

                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("invalid file type. -> " + zipPath);
                    }
                }
            }
            

        }
    }
}
