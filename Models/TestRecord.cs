using System;

namespace ZebraPrinterMonitor.Models
{
    public class TestRecord
    {
        public string? TR_SerialNum { get; set; }
        public string? TR_ID { get; set; }
        public DateTime? TR_DateTime { get; set; }
        public decimal? TR_Isc { get; set; }
        public decimal? TR_Voc { get; set; }
        public decimal? TR_Pm { get; set; }
        public decimal? TR_Ipm { get; set; }
        public decimal? TR_Vpm { get; set; }
        public decimal? TR_CellEfficiency { get; set; }
        public decimal? TR_FF { get; set; }
        public string? TR_Grade { get; set; }
        public decimal? TR_Temp { get; set; }
        public decimal? TR_Irradiance { get; set; }
        public decimal? TR_Rs { get; set; }
        public decimal? TR_Rsh { get; set; }
        public string? TR_CellArea { get; set; }
        public string? TR_Operater { get; set; }
        public string? TR_FontColor { get; set; }
        public string? TR_BackColor { get; set; }
        public int? TR_Print { get; set; } = 0;

        public string GetDisplayText()
        {
            return $"序列号: {TR_SerialNum ?? "N/A"} - 日期: {TR_DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A"}";
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

        public string FormatNumber(decimal? value)
        {
            if (!value.HasValue) return "N/A";
            
            // 如果值很小，使用科学记数法
            if (Math.Abs(value.Value) < 0.001m && value.Value != 0)
            {
                return value.Value.ToString("E2");
            }
            
            // 否则使用固定小数点
            return value.Value.ToString("F3");
        }

        public override string ToString()
        {
            return $"TestRecord: SerialNum={TR_SerialNum}, ID={TR_ID}";
        }
    }
} 