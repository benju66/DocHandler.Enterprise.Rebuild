using System.Threading.Tasks;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Interface for configuration export and import operations
    /// </summary>
    public interface IConfigurationExportImportService
    {
        /// <summary>
        /// Export current configuration to YAML file
        /// </summary>
        Task<string> ExportConfigurationAsync(string? filePath = null, ExportOptions? options = null);

        /// <summary>
        /// Import configuration from YAML file with validation
        /// </summary>
        Task<ImportResult> ImportConfigurationAsync(string filePath, ImportOptions? options = null);

        /// <summary>
        /// Create a backup of the current configuration
        /// </summary>
        Task<string> CreateConfigurationBackup();

        /// <summary>
        /// Validate configuration structure and values
        /// </summary>
        ConfigurationValidationResult ValidateConfiguration(HierarchicalAppConfiguration config);
    }
} 