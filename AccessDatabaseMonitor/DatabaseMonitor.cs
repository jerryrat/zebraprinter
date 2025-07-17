using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace AccessDatabaseMonitor
{
    public class DatabaseMonitor
    {
        private readonly string _connectionString;
        private readonly System.Threading.Timer _monitorTimer;
        private readonly HashSet<TestRecord> _knownRecords;
        private bool _isRunning;

        public event Action<List<TestRecord>>? NewRecordsDetected;
        public event Action<string>? ErrorOccurred;

        public DatabaseMonitor(string databasePath)
        {
            // 使用支持并发访问的连接字符串
            _connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Mode=Share Deny None;Persist Security Info=false;Jet OLEDB:Database Locking Mode=1;";
            _knownRecords = new HashSet<TestRecord>();
            _monitorTimer = new System.Threading.Timer(CheckForNewRecords, null, Timeout.Infinite, Timeout.Infinite);
        }

        public async Task<bool> TestConnectionAsync()
        {
            try
            {
                using var connection = new OleDbConnection(_connectionString);
                await connection.OpenAsync();
                return true;
            }
            catch (Exception ex)
            {
                ErrorOccurred?.Invoke($"Database connection test failed: {ex.Message}");
                return false;
            }
        }

        public async Task InitializeAsync()
        {
            try
            {
                var currentRecords = await GetAllRecordsAsync();
                _knownRecords.Clear();
                foreach (var record in currentRecords)
                {
                    _knownRecords.Add(record);
                }
            }
            catch (Exception ex)
            {
                ErrorOccurred?.Invoke($"Database initialization failed: {ex.Message}");
            }
        }

        public void StartMonitoring(int intervalSeconds = 5)
        {
            if (!_isRunning)
            {
                _isRunning = true;
                _monitorTimer.Change(TimeSpan.Zero, TimeSpan.FromSeconds(intervalSeconds));
            }
        }

        public void StopMonitoring()
        {
            if (_isRunning)
            {
                _isRunning = false;
                _monitorTimer.Change(Timeout.Infinite, Timeout.Infinite);
            }
        }

        private async void CheckForNewRecords(object? state)
        {
            if (!_isRunning) return;

            try
            {
                var currentRecords = await GetAllRecordsAsync();
                var newRecords = currentRecords.Where(record => !_knownRecords.Contains(record)).ToList();

                if (newRecords.Any())
                {
                    foreach (var record in newRecords)
                    {
                        _knownRecords.Add(record);
                    }
                    NewRecordsDetected?.Invoke(newRecords);
                }
            }
            catch (Exception ex)
            {
                ErrorOccurred?.Invoke($"Monitoring error: {ex.Message}");
            }
        }

        private async Task<List<TestRecord>> GetAllRecordsAsync()
        {
            var records = new List<TestRecord>();

            using var connection = new OleDbConnection(_connectionString);
            await connection.OpenAsync();

            using var command = new OleDbCommand("SELECT TR_SerialNum, TR_ID FROM TestRecord", connection);
            using var reader = await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                var serialNum = reader.IsDBNull("TR_SerialNum") ? string.Empty : reader.GetString("TR_SerialNum");
                var id = reader.IsDBNull("TR_ID") ? string.Empty : reader.GetString("TR_ID");

                records.Add(new TestRecord(serialNum, id));
            }

            return records;
        }

        public async Task<List<TestRecord>> GetRecentRecordsAsync(int count = 10)
        {
            var records = new List<TestRecord>();

            try
            {
                using var connection = new OleDbConnection(_connectionString);
                await connection.OpenAsync();

                using var command = new OleDbCommand($"SELECT TOP {count} TR_SerialNum, TR_ID FROM TestRecord ORDER BY TR_SerialNum DESC", connection);
                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    var serialNum = reader.IsDBNull("TR_SerialNum") ? string.Empty : reader.GetString("TR_SerialNum");
                    var id = reader.IsDBNull("TR_ID") ? string.Empty : reader.GetString("TR_ID");

                    records.Add(new TestRecord(serialNum, id));
                }
            }
            catch (Exception ex)
            {
                ErrorOccurred?.Invoke($"Failed to get recent records: {ex.Message}");
            }

            return records;
        }

        public void Dispose()
        {
            StopMonitoring();
            _monitorTimer?.Dispose();
        }
    }
} 