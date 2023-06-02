using System;
using System.IO;

namespace Nuss_MS_EMFSPOOL_To_MS_EMF
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                Console.WriteLine();
                Console.WriteLine("Nusstudios' MS-EMFSPOOL to standard MS-EMF tool v4.2.0.0");
                Console.WriteLine();
                Console.WriteLine("Working...");
                processSpl(args[0], Convert.ToBoolean(args[1]), Convert.ToBoolean(args[2]), Convert.ToInt32(args[3]), Convert.ToBoolean(args[4]), Convert.ToBoolean(args[5]));
                Console.WriteLine();
                Console.WriteLine("Finished successfully");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("Nusstudios' MS-EMFSPOOL to standard MS-EMF tool v4.2");
                Console.WriteLine();
                Console.Write("Enter file name (if not needed, enter .SPL): ");
                string fnStr = Console.ReadLine();
                Console.WriteLine();
                Console.Write("Process multiple files? [true/false]: ");
                string procMultiF = Console.ReadLine();
                Console.WriteLine();
                string enumByIndex = "true";
                string i = "0";

                if (Convert.ToBoolean(procMultiF))
                {
                    Console.Write("Enumerate by index (default) or extension (.SPL)? [true/false]: ");
                    enumByIndex = Console.ReadLine();
                    Console.WriteLine();

                    if (Convert.ToBoolean(enumByIndex))
                    {
                        Console.Write("MS-EMFSPOOL start index [any integer value]: ");
                        i = Console.ReadLine();
                        Console.WriteLine();
                    }
                }

                Console.Write("Unpack all images or keep first image only (per spl emf file)? [true/false]: ");
                string allP = Console.ReadLine();

                string setFI = "true";

                if (Convert.ToBoolean(procMultiF))
                {
                    Console.WriteLine();
                    Console.Write("Set first page index, or leave it as it is? [true/false]: ");
                    setFI = Console.ReadLine();
                }

                Console.WriteLine();
                Console.WriteLine("Working...");
                processSpl(fnStr, Convert.ToBoolean(procMultiF), Convert.ToBoolean(enumByIndex), Convert.ToInt32(i), Convert.ToBoolean(allP), Convert.ToBoolean(setFI));
            }
        }

        public enum RecordTye
        {
            EMRI_METAFILE = 0x00000001,
            EMRI_ENGINE_FONT = 0x00000002,
            EMRI_DEVMODE = 0x00000003,
            EMRI_TYPE1_FONT = 0x00000004,
            EMRI_PRESTARTPAGE = 0x00000005,
            EMRI_DESIGNVECTOR = 0x00000006,
            EMRI_SUBSET_FONT = 0x00000007,
            EMRI_DELTA_FONT = 0x00000008,
            EMRI_FORM_METAFILE = 0x00000009,
            EMRI_BW_METAFILE = 0x0000000A,
            EMRI_BW_FORM_METAFILE = 0x0000000B,
            EMRI_METAFILE_DATA = 0x0000000C,
            EMRI_METAFILE_EXT = 0x0000000D,
            EMRI_BW_METAFILE_EXT = 0x0000000E,
            EMRI_ENGINE_FONT_EXT = 0x0000000F,
            EMRI_TYPE1_FONT_EXT = 0x00000010,
            EMRI_DESIGNVECTOR_EXT = 0x00000011,
            EMRI_SUBSET_FONT_EXT = 0x00000012,
            EMRI_DELTA_FONT_EXT = 0x00000013,
            EMRI_PS_JOB_DATA = 0x00000014,
            EMRI_EMBED_FONT_EXT = 0x00000015
        }

        private static void processSpl(string fnStr, bool procMultiF, bool enumByIndex, int i, bool allP, bool setFI)
        {
            string filename = fnStr;
            int index = i;

            string fn = "";

            if (procMultiF)
            {
                if (enumByIndex)
                {
                    string[] files = Directory.GetFiles(Environment.CurrentDirectory);

                    foreach (string file in files)
                    {
                        if (filename == ".SPL")
                        {
                            if (file.IndexOf(".SPL") != -1)
                            {
                                int outNum = 0;
                                string fileNoPath = Path.GetFileName(file);
                                string fileNoExt = fileNoPath.Substring(0, fileNoPath.IndexOf("."));

                                if (Int32.TryParse(fileNoExt, out outNum))
                                {
                                    string newPath = Environment.CurrentDirectory + "\\" + outNum.ToString() + ".SPL";
                                    File.Move(file, newPath);
                                }
                            }
                        }
                        else
                        {
                            string fileNoPath = Path.GetFileName(file);
                            string fnNoExt = filename.Substring(0, filename.IndexOf("."));

                            if (fileNoPath.IndexOf(fnNoExt) != -1)
                            {
                                string numAndExt = fileNoPath.Replace(fnNoExt, "");
                                string num = numAndExt.Substring(0, numAndExt.IndexOf("."));
                                int outNum = 0;

                                if (Int32.TryParse(num, out outNum))
                                {
                                    string fnExt = filename.Substring(filename.IndexOf("."));
                                    string newPath = Environment.CurrentDirectory + "\\" + fnNoExt + outNum + fnExt;
                                    File.Move(file, newPath);
                                }
                            }
                        }
                    }

                    while (File.Exists((fn = Environment.CurrentDirectory + "\\" + (filename != "" ? filename.Substring(0, filename.LastIndexOf(".")) : "") + index.ToString() + filename.Substring(filename.LastIndexOf(".")))))
                    {
                        msEmfSpoolToMsEmf(fn, allP, setFI);
                        index++;
                    }
                }
                else
                {
                    string[] files = Directory.GetFiles(Environment.CurrentDirectory);
                    string prevFile = "";

                    foreach (string file in files)
                    {
                        if (Path.GetExtension(file) == ".SPL" || Path.GetExtension(file) == ".spl")
                        {
                            msEmfSpoolToMsEmf(file, allP, setFI);
                            prevFile = file;
                        }
                    }
                }
            }
            else
            {
                msEmfSpoolToMsEmf(filename, allP, setFI);
            }
        }

        private static void msEmfSpoolToMsEmf(string fnStr, bool allP, bool setFI)
        {
            bool wasBroken = false;
            int imagenum = 0;
            FileStream splEmfStream = new FileStream(fnStr, FileMode.Open);
            BinaryReader splEmfStreamReader = new BinaryReader(splEmfStream);
            UInt32 dwVersion = splEmfStreamReader.ReadUInt32();

            // dwVersion must equal to 0x00010000
            if (dwVersion == 0x00010000)
            {
                /* In EMF Spool format, every record's first two unsigned 32 bit integers
                should be the record type, and the cjSize*/
                UInt32 cjSize = splEmfStreamReader.ReadUInt32();
                // Seek back
                splEmfStream.Seek(0, SeekOrigin.Begin);
                splEmfStream.Seek(cjSize, SeekOrigin.Begin);

                // The following loop aims to hook only pure, standard EMF files of format MS-EMF from format MS-EMFSPOOL
                while (splEmfStream.Position != splEmfStream.Length)
                {
                    // Record types after emf spool header are called 'data records'
                    UInt32 dataRecord = splEmfStreamReader.ReadUInt32();

                    if (dataRecord == (UInt32)RecordTye.EMRI_METAFILE_DATA || dataRecord == (UInt32)RecordTye.EMRI_METAFILE || dataRecord == (UInt32)RecordTye.EMRI_BW_METAFILE || dataRecord == (UInt32)RecordTye.EMRI_FORM_METAFILE || dataRecord == (UInt32)RecordTye.EMRI_BW_FORM_METAFILE)
                    {
                        // Process record if it's a page content record
                        cjSize = splEmfStreamReader.ReadUInt32();

                        if (cjSize == 0)
                        {
                            /* In case of page content records,
                            if cjSize is 0, it means there are
                            no more pages left.*/
                            wasBroken = true;
                            Console.WriteLine();
                            Console.WriteLine("Finished successfully");
                            splEmfStreamReader.Close();
                            splEmfStream.Close();
                            break;
                        }

                        byte[] emfBuffer = splEmfStreamReader.ReadBytes((Int32)cjSize);

                        if (imagenum >= 1)
                        {
                            if (setFI && imagenum == 1)
                            {
                                File.Move(Environment.CurrentDirectory + "\\" + Path.GetFileNameWithoutExtension(fnStr) + ".emf", Path.GetFileNameWithoutExtension(fnStr) + "_0.emf");
                            }

                            File.WriteAllBytes(Environment.CurrentDirectory + "\\" + Path.GetFileNameWithoutExtension(fnStr) + "_" + imagenum.ToString() + ".emf", emfBuffer);
                        }
                        else
                        {
                            File.WriteAllBytes(Environment.CurrentDirectory + "\\" + Path.GetFileNameWithoutExtension(fnStr) + ".emf", emfBuffer);
                        }

                        if (!allP)
                        {
                            splEmfStreamReader.Close();
                            splEmfStream.Close();
                            break;
                        }

                        imagenum++;
                    }
                    else if (dataRecord == (UInt32)RecordTye.EMRI_METAFILE_EXT || dataRecord == (UInt32)RecordTye.EMRI_BW_METAFILE_EXT)
                    {
                        // Page offset records are ignored
                        cjSize = splEmfStreamReader.ReadUInt32();
                        splEmfStream.Seek(splEmfStream.Position + cjSize, SeekOrigin.Begin);
                    }
                    else if (dataRecord == (UInt32)RecordTye.EMRI_ENGINE_FONT || dataRecord == (UInt32)RecordTye.EMRI_TYPE1_FONT || dataRecord == (UInt32)RecordTye.EMRI_DESIGNVECTOR || dataRecord == (UInt32)RecordTye.EMRI_SUBSET_FONT || dataRecord == (UInt32)RecordTye.EMRI_DELTA_FONT)
                    {
                        // Font records are ignored
                        cjSize = splEmfStreamReader.ReadUInt32();
                        splEmfStream.Seek(splEmfStream.Position + cjSize, SeekOrigin.Begin);
                    }
                    else if (dataRecord == (UInt32)RecordTye.EMRI_ENGINE_FONT_EXT || dataRecord == (UInt32)RecordTye.EMRI_TYPE1_FONT_EXT || dataRecord == (UInt32)RecordTye.EMRI_DESIGNVECTOR_EXT || dataRecord == (UInt32)RecordTye.EMRI_SUBSET_FONT_EXT || dataRecord == (UInt32)RecordTye.EMRI_DELTA_FONT_EXT)
                    {
                        // Font offset records are ignored
                        cjSize = splEmfStreamReader.ReadUInt32();
                        splEmfStream.Seek(splEmfStream.Position + cjSize, SeekOrigin.Begin);
                    }
                    else if (dataRecord == (UInt32)RecordTye.EMRI_DEVMODE)
                    {
                        // EMRI_DEVMODE record is ignored
                        cjSize = splEmfStreamReader.ReadUInt32();
                        splEmfStream.Seek(splEmfStream.Position + cjSize, SeekOrigin.Begin);
                    }
                    else if (dataRecord == (UInt32)RecordTye.EMRI_PRESTARTPAGE)
                    {
                        // EMRI_PRESTARTPAGE record is ignored
                        cjSize = splEmfStreamReader.ReadUInt32();
                        splEmfStream.Seek(splEmfStream.Position + cjSize, SeekOrigin.Begin);
                    }
                    else if (dataRecord == (UInt32)RecordTye.EMRI_PS_JOB_DATA)
                    {
                        // EMRI_PS_JOB_DATA record is ignored
                        cjSize = splEmfStreamReader.ReadUInt32();
                        splEmfStream.Seek(splEmfStream.Position + cjSize, SeekOrigin.Begin);
                    }
                    else
                    {
                        // UNKNOWN record is ignored
                        cjSize = splEmfStreamReader.ReadUInt32();
                        splEmfStream.Seek(splEmfStream.Position + cjSize, SeekOrigin.Begin);
                    }
                }

                if (!wasBroken)
                {
                    Console.WriteLine();
                    Console.WriteLine("Finished successfully");
                }
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("'" + fnStr + "' is not a valid MS-EMFSPOOL file");
            }
        }
    }
}
