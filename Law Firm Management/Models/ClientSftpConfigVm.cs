namespace Law_Firm_Management.Models
{
    public class ClientSftpConfigVm
    {
        public int ClientSftpConfigId { get; set; }
        public int ClientId { get; set; }
        public string? ConfigName { get; set; }
        public string? Host { get; set; }
        public int Port { get; set; } = 22;
        public string? Username { get; set; }
        public string? AuthType { get; set; }
        public string? PasswordValue { get; set; }
        public string? KeyFileName { get; set; }
        public string? RemoteInboundPath { get; set; }
        public string? RemoteOutboundPath { get; set; }
        public string? ArchivePath { get; set; }
        public string? ErrorPath { get; set; }
        public bool IsActive { get; set; }
        public string? Notes { get; set; }
        public DateTime? LastConnectionTestDate { get; set; }
        public string? LastConnectionTestStatus { get; set; }
    }
}