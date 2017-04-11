using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace ExcelDemo.Models
{
    public sealed class FileInfo 
    {
        public string Name { get; set; }
        public long Length { get; set; }
        public string DirectoryName { get; set; }
        public string Directory { get; set; }
        public bool IsReadOnly { get; set; }
        public bool Exists { get; set; }
        public string FullName { get; set; }
        public string Extension{ get; set; }
        public DateTime CreationTime { get; set; }
        public DateTime CreationTimeUtc { get; set; }
        public DateTime LastAccessTime { get; set; }
        public DateTime LastAccessTimeUtc { get; set; }
        public DateTime LastWriteTime { get; set; }
        public DateTime LastWriteTimeUtc { get; set;  }
        public string Attributes {get;set;}
    }
}