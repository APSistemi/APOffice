using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.BZip2;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Zip.Compression;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using ICSharpCode.SharpZipLib.GZip;

class clsSharpZip
{
    /* 
    strPathZip = File .zip da creare          
    strFile    = file da includere nello zip
    strWhere   = path dove si trova il file da zippare per evitare la path nello zip
    */
    public void Zippa( string strPathZip, string[] strFils, string strWhere)
    {

        ZipOutputStream sZip = new ZipOutputStream(File.Create(strPathZip));

        //Directory.SetCurrentDirectory("C:\\TEMP");
        Directory.SetCurrentDirectory(strWhere);

        foreach (string sFil in strFils)
        {

            //apro un file stream sul file
            FileStream fs = File.OpenRead(sFil);
            //creo l'array di byte
            byte[] buffer = new byte[fs.Length];
            //Leggo il buffer
            fs.Read(buffer, 0, buffer.Length);

            //creo una nuova entry
            ZipEntry input = new ZipEntry(sFil);
            //Aggiungo la entry al file zip
            sZip.PutNextEntry(input);
            //Scrivo il buffer
            sZip.Write(buffer, 0, buffer.Length);
            fs.Close();
            fs.Dispose();

        }

        sZip.Finish();
        sZip.Close();

    }

    public void UnZip(string strPathZip, string strPathDest, Boolean bolRenData)
    {

        FileInfo fInfo = new FileInfo(strPathZip);

        if (fInfo.Length > 0)
        {

            try
            {
                using (ZipInputStream s = new ZipInputStream(File.OpenRead(strPathZip)))
                {

                    ZipEntry theEntry;

                    while ((theEntry = s.GetNextEntry()) != null)
                    {

                        //Console.WriteLine(theEntry.Name);

                        string directoryName = Path.GetDirectoryName(theEntry.Name);
                        string fileName = Path.GetFileName(theEntry.Name);

                        if (bolRenData)
                            fileName += "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetExtension(fileName);

                        //// create directory
                        //if (directoryName.Length > 0)
                        //{
                        //    //Directory.CreateDirectory(directoryName);
                        //}

                        if (fileName != String.Empty)
                        {
                            //using (FileStream streamWriter = File.Create(theEntry.Name))
                            using (FileStream streamWriter = File.Create(strPathDest + "\\" + fileName))
                            {

                                int size = 2048;
                                byte[] data = new byte[2048];
                                while (true)
                                {
                                    size = s.Read(data, 0, data.Length);
                                    if (size > 0)
                                    {
                                        streamWriter.Write(data, 0, size);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }

                }

            }

            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);

                string s2 = Path.GetDirectoryName(strPathZip);
                s2 += "\\Err";
                s2 += Path.GetFileName(strPathZip);

                File.Move(strPathZip, s2);


            }

        }
    }

    public ArrayList UnVedi(string strPathZip)
    {

        FileInfo fInfo = new FileInfo(strPathZip);

        byte[] data = new byte[4096];

        ArrayList ar = new ArrayList();

        if (fInfo.Length > 0)
        {

            using (ZipInputStream s = new ZipInputStream(File.OpenRead(strPathZip)))
            {

                ZipEntry theEntry;
                while ((theEntry = s.GetNextEntry()) != null)
                {
                    //Console.WriteLine("Name : {0}", theEntry.Name);
                    //Console.WriteLine("Date : {0}", theEntry.DateTime);
                    //Console.WriteLine("Size : (-1, if the size information is in the footer)");
                    //Console.WriteLine("      Uncompressed : {0}", theEntry.Size);
                    //Console.WriteLine("      Compressed   : {0}", theEntry.CompressedSize);

                    ar.Add(theEntry.Name);

                    //if (theEntry.IsFile)
                    //{

                    //    // Assuming the contents are text may be ok depending on what you are doing
                    //    // here its fine as its shows how data can be read from a Zip archive.
                    //    Console.Write("Show entry text (y/n) ?");

                    //    if (Console.ReadLine() == "y")
                    //    {
                    //        int size = s.Read(data, 0, data.Length);
                    //        while (size > 0)
                    //        {
                    //            Console.Write(Encoding.ASCII.GetString(data, 0, size));
                    //            size = s.Read(data, 0, data.Length);
                    //        }
                    //    }
                    //    Console.WriteLine();
                    //}

                }

                // Close can be ommitted as the using statement will do it automatically
                // but leaving it here reminds you that is should be done.
                s.Close();
            }
        }

        return ar;

    }
}
