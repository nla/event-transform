using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Xsl;

namespace EventTransform
{
    class Launcher
    {
        static void Main(string[] args)
        {

            if (args.Length == 4)
            {
                Processor p = new Processor(args);
                p.ActionState();
            }
            else
            {
                Console.WriteLine("Parameters required");
                Console.WriteLine("- Event log file");
                Console.WriteLine("- XSLT file");
                Console.WriteLine("- Output file");
                Console.WriteLine("- Log filename");
            }

        }

    }

    class Processor : IDisposable
    {

        enum ConfigItem
        {
            FileEvtx,
            FileXslt,
            FileOutput,
            FileLog,
            State
        }

        enum Action
        {
            InitializeLogging,
            CheckFiles, CheckFileError,
            PrepareFiles,
            ProcessEvents,
            EndRun
        }

        private Action whatDo;
        private Dictionary<ConfigItem, string> config;

        private EventLogReader eventReader;

        private XmlWriter xmlWriter;
        private XslCompiledTransform xmlTransform;
        private XmlDocument xmlDoc;
        private StringBuilder outBuffer;

        private int lastReadEventCount;

        StreamWriter outFileWriter, logFileWriter;

        public Processor(string[] args)
        {
            config = new Dictionary<ConfigItem, string>
            {
                { ConfigItem.FileEvtx, args[0] },
                { ConfigItem.FileXslt, args[1] },
                { ConfigItem.FileOutput, args[2] },
                { ConfigItem.FileLog, args[3] }
            };

            whatDo = Action.InitializeLogging;
            lastReadEventCount = 0;

            xmlTransform = new XslCompiledTransform();
            xmlDoc = new XmlDocument();
            outBuffer = new StringBuilder();

            xmlWriter = XmlWriter.Create(outBuffer, new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                ConformanceLevel = ConformanceLevel.Fragment
            });

        }

        public void Dispose()
        {
            eventReader.Dispose();

            if (outFileWriter != null)
            {
                logFileWriter.WriteLine("Closing output file");
                outFileWriter.Close();
                outFileWriter.Dispose();
            }

            if (logFileWriter != null)
            {
                logFileWriter.WriteLine("Closing log file");
                logFileWriter.Close();
                logFileWriter.Dispose();
            }

            // This object will be cleaned up by the Dispose method.
            // Therefore, you should call GC.SupressFinalize to
            // take this object off the finalization queue
            // and prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }

        public void ActionState()
        {
            while (whatDo != Action.EndRun)
            {
                switch (whatDo)
                {
                    case Action.InitializeLogging:
                        try
                        {
                            logFileWriter = new StreamWriter(config[ConfigItem.FileLog], false)
                            {
                                AutoFlush = true
                            };

                            logFileWriter.WriteLine("Event file: " + config[ConfigItem.FileEvtx]);
                            logFileWriter.WriteLine("XSLT file: " + config[ConfigItem.FileXslt]);
                            logFileWriter.WriteLine("Log file: " + config[ConfigItem.FileLog]);
                            logFileWriter.WriteLine("Output file: " + config[ConfigItem.FileOutput]);

                            whatDo = Action.CheckFiles;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Error initializing log file");
                            Console.WriteLine(e.Message);
                            whatDo = Action.EndRun;
                            break;
                        }
                        break;

                    case Action.CheckFiles:
                        logFileWriter.WriteLine("Checking files ...");
                        whatDo =
                            (CanReadFile(config[ConfigItem.FileEvtx]) &&
                            CanReadFile(config[ConfigItem.FileXslt]) &&
                            CanWriteFile(config[ConfigItem.FileOutput]))
                            ? Action.PrepareFiles : Action.CheckFileError;
                        break;

                    case Action.CheckFileError:
                        logFileWriter.WriteLine("Error checking files. Please check the following files and necessary access.");
                        whatDo = Action.EndRun;
                        break;

                    case Action.PrepareFiles:
                        logFileWriter.WriteLine("Preparing files ...");

                        // load xslt 

                        try
                        {
                            xmlTransform.Load(config[ConfigItem.FileXslt]);
                        }
                        catch (Exception e)
                        {
                            logFileWriter.WriteLine("Error loading XSLT");
                            logFileWriter.WriteLine(e.ToString());
                            whatDo = Action.EndRun;
                        }

                        try
                        {
                            eventReader = new EventLogReader(config[ConfigItem.FileEvtx], PathType.FilePath);
                        }
                        catch (Exception e)
                        {
                            logFileWriter.WriteLine("Error reading event log file");
                            logFileWriter.WriteLine(e.ToString());
                            whatDo = Action.EndRun;
                        }

                        try
                        {
                            outFileWriter = new StreamWriter(config[ConfigItem.FileOutput], false)
                            {
                                AutoFlush = true
                            };
                        }
                        catch (Exception e)
                        {
                            logFileWriter.WriteLine("Error oepning output file for writing");
                            logFileWriter.WriteLine(e.ToString());
                            whatDo = Action.EndRun;
                        }

                        if (whatDo != Action.EndRun)
                        {
                            whatDo = Action.ProcessEvents;
                        }

                        break;

                    case Action.ProcessEvents:
                        logFileWriter.WriteLine("Reading and processing events ...");

                        try
                        {
                            EventRecord currentEvent;

                            while ((currentEvent = eventReader.ReadEvent()) != null)
                            {
                                lastReadEventCount++;
                                // transform to xml, apply xsl
                                outBuffer.Clear();
                                xmlDoc.LoadXml(currentEvent.ToXml());
                                xmlTransform.Transform(xmlDoc, null, xmlWriter);
                                outFileWriter.WriteLine(outBuffer.ToString());
                            }

                            logFileWriter.WriteLine("Processed " + lastReadEventCount + " records");
                            whatDo = Action.EndRun;
                        }
                        catch (Exception e)
                        {
                            logFileWriter.WriteLine("Last successfully read record # " + lastReadEventCount);
                            logFileWriter.WriteLine(e.ToString());
                            whatDo = Action.EndRun;
                        }

                        break;

                    default:
                        logFileWriter.WriteLine("Unhandled action state. Complain to your programmer. :-)");
                        whatDo = Action.EndRun;
                        break;
                }
            }

        }

        private bool CanReadFile(string filename)
        {
            bool success = File.Exists(filename);

            if (success)
            {
                try
                {
                    FileInfo finfo = new FileInfo(filename);
                    if (finfo.Length > 0)
                    {
                        FileStream fs = File.OpenRead(filename);
                        fs.Close();
                        success = true;
                    }
                    else
                    {
                        // return false for zero length files 
                        success = false;
                    }
                }
                catch
                {
                    // return false on error getting file info or calling OpenRead()
                    success = false;
                }
            }

            return success;
        }

        private bool CanWriteFile(string filename)
        {
            bool success = false;

            try
            {
                FileStream fs = File.OpenWrite(filename);
                fs.Close();
                success = true;
            }
            catch
            {
                // return false on error opening file for writing
                success = false;
            }

            return success;
        }

    }
}
