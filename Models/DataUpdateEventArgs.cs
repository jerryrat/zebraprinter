using System;
using System.Collections.Generic;

namespace ZebraPrinterMonitor.Models
{
    /// <summary>
    /// 🔧 统一数据更新事件参数 - 基于GetLastRecord的完整数据刷新
    /// </summary>
    public class DataUpdateEventArgs : EventArgs
    {
        /// <summary>
        /// 触发更新的最后记录
        /// </summary>
        public TestRecord LastRecord { get; set; }
        
        /// <summary>
        /// 最新的记录列表（通常50条）
        /// </summary>
        public List<TestRecord> RecentRecords { get; set; }
        
        /// <summary>
        /// 更新类型描述
        /// </summary>
        public string UpdateType { get; set; }
        
        /// <summary>
        /// 变化详情
        /// </summary>
        public string ChangeDetails { get; set; }

        public DataUpdateEventArgs(TestRecord lastRecord, List<TestRecord> recentRecords, string updateType, string changeDetails)
        {
            LastRecord = lastRecord;
            RecentRecords = recentRecords;
            UpdateType = updateType;
            ChangeDetails = changeDetails;
        }
    }
} 