using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLEConsole
{
    internal class FileWriter
    {
        private FileStream fileStream;
        public string CreateFile(string path, string filenameTemplate, int id)
        {
            // Generate the full filename
            string filename = String.Format(filenameTemplate, id);
            string fullPath = Path.Combine(path, filename);

            // Check if the file already exists
            if (File.Exists(fullPath))
            {
                throw new IOException("File already exists: " + fullPath);
            }

            // Open the file for writing
            fileStream = new FileStream(fullPath, FileMode.CreateNew, FileAccess.Write);
            return fullPath;
        }

        public void WriteData(string data)
        {
            if (fileStream == null)
            {
                throw new InvalidOperationException("File not created.");
            }

            using (StreamWriter writer = new StreamWriter(fileStream, encoding: Encoding.UTF8, bufferSize: 1024,leaveOpen: true))
            {
                writer.WriteLine(data);
                writer.Flush();
            }
        }

        public void WriteData(byte[] data)
        {
            if (fileStream == null)
            {
                throw new InvalidOperationException("File not created.");
            }
            fileStream.Write(data, 0, data.Length); 
            fileStream.Flush();
        }
        public void CloseFile()
        {
            if (fileStream != null)
            {
                fileStream.Close();
                fileStream = null;
            }
        }

        // Ensure the FileStream is properly disposed when the object is destroyed
        ~FileWriter()
        {
            fileStream?.Dispose();
        }
    }
}
