using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;
using ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal sealed partial class ZipWorker : IDisposable
    {
        private const string FileSharedStrings = "xl/sharedStrings.{0}";
        private const string FileStyles = "xl/styles.{0}";
        private const string FileWorkbook = "xl/workbook.{0}";
        private const string FileRels = "xl/_rels/workbook.{0}.rels";

        private const string Format = "xml";
        private const string BinFormat = "bin";

        private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings 
        {
            IgnoreComments = true, 
            IgnoreWhitespace = true,
        };

        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private bool _disposed;
        private Stream _zipStream;
        private ZipArchive _zipFile;

        /// <summary>
        /// Initializes a new instance of the <see cref="ZipWorker"/> class. 
        /// </summary>
        /// <param name="fileStream">The zip file stream.</param>
        public ZipWorker(Stream fileStream)
        {
            _zipStream = fileStream ?? throw new ArgumentNullException(nameof(fileStream));
            _zipFile = new ZipArchive(fileStream);
            _entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in _zipFile.Entries)
            {
                _entries.Add(entry.FullName.Replace('\\', '/'), entry);
            }
        }

        /// <summary>
        /// Gets the shared strings reader.
        /// </summary>
        public RecordReader GetSharedStringsReader(XmlProperNamespaces properNamespaces)
        {
            var entry = FindEntry(string.Format(CultureInfo.InvariantCulture, FileSharedStrings, Format));
            if (entry != null)
                return new XmlSharedStringsReader(XmlReader.Create(entry.Open(), XmlSettings), properNamespaces);

            entry = FindEntry(string.Format(CultureInfo.InvariantCulture, FileSharedStrings, BinFormat));
            if (entry != null)
                return new BiffSharedStringsReader(entry.Open());

            return null;
        }

        /// <summary>
        /// Gets the styles reader.
        /// </summary>
        public RecordReader GetStylesReader(XmlProperNamespaces properNamespaces)
        {
            var entry = FindEntry(string.Format(CultureInfo.InvariantCulture, FileStyles, Format));
            if (entry != null)
                return new XmlStylesReader(XmlReader.Create(entry.Open(), XmlSettings), properNamespaces);

            entry = FindEntry(string.Format(CultureInfo.InvariantCulture, FileStyles, BinFormat));
            if (entry != null)
                return new BiffStylesReader(entry.Open());

            return null;
        }

        /// <summary>
        /// Gets the workbook reader.
        /// </summary>
        public RecordReader GetWorkbookReader(XmlProperNamespaces properNamespaces)
        {
            var entry = FindEntry(string.Format(CultureInfo.InvariantCulture, FileWorkbook, Format));
            if (entry != null)
                return new XmlWorkbookReader(XmlReader.Create(entry.Open(), XmlSettings), properNamespaces);

            entry = FindEntry(string.Format(CultureInfo.InvariantCulture, FileWorkbook, BinFormat));
            if (entry != null)
                return new BiffWorkbookReader(entry.Open());

            throw new Exceptions.HeaderException(Errors.ErrorZipNoOpenXml);
        }

        public RecordReader GetWorksheetReader(string sheetPath, XmlProperNamespaces properNamespaces)
        {
            // its possible sheetPath starts with /xl. in this case trim the /
            // see the test "Issue_11522_OpenXml"
            if (sheetPath.StartsWith("/xl/", StringComparison.OrdinalIgnoreCase))
                sheetPath = sheetPath.Substring(1);
            else
                sheetPath = "xl/" + sheetPath;

            var zipEntry = FindEntry(sheetPath);
            if (zipEntry != null)
            {
                return Path.GetExtension(sheetPath) switch
                {
                    ".xml" => new XmlWorksheetReader(XmlReader.Create(zipEntry.Open(), XmlSettings), properNamespaces),
                    ".bin" => new BiffWorksheetReader(zipEntry.Open()),
                    _ => null,
                };
            }

            return null;
        }

        /// <summary>
        /// Gets the workbook rels stream.
        /// </summary>
        /// <returns>The rels stream.</returns>
        public Stream GetWorkbookRelsStream()
        {
            var zipEntry = FindEntry(string.Format(CultureInfo.InvariantCulture, FileRels, Format));
            if (zipEntry != null)
                return zipEntry.Open();

            zipEntry = FindEntry(string.Format(CultureInfo.InvariantCulture, FileRels, BinFormat));
            if (zipEntry != null)
                return zipEntry.Open();

            return null;
        }

        private ZipArchiveEntry FindEntry(string name)
        {
            if (_entries.TryGetValue(name, out var entry))
                return entry;
            return null;
        }
    }

    internal partial class ZipWorker
    {
        ~ZipWorker()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_zipFile != null)
                    {
                        _zipFile.Dispose();
                        _zipFile = null;
                    }

                    if (_zipStream != null)
                    {
                        _zipStream.Dispose();
                        _zipStream = null;
                    }
                }

                _disposed = true;
            }
        }
    }
}