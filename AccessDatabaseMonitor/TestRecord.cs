using System;

namespace AccessDatabaseMonitor
{
    public class TestRecord
    {
        public string TR_SerialNum { get; set; } = string.Empty;
        public string TR_ID { get; set; } = string.Empty;
        public DateTime LastModified { get; set; }

        public TestRecord()
        {
            LastModified = DateTime.Now;
        }

        public TestRecord(string serialNum, string id)
        {
            TR_SerialNum = serialNum;
            TR_ID = id;
            LastModified = DateTime.Now;
        }

        public override string ToString()
        {
            return $"SerialNum: {TR_SerialNum}, ID: {TR_ID}, Modified: {LastModified:yyyy-MM-dd HH:mm:ss}";
        }

        public override bool Equals(object? obj)
        {
            if (obj is TestRecord other)
            {
                return TR_SerialNum == other.TR_SerialNum && TR_ID == other.TR_ID;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(TR_SerialNum, TR_ID);
        }
    }
} 