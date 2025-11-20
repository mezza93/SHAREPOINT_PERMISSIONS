# Changelog

All notable changes to the SharePoint Permissions Audit Tool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] - 2025-11-20

### Added
- Configuration file support (JSON format) for easier recurring audits
- Retry logic with exponential backoff for transient failures
- Progress tracking with Write-Progress cmdlets
- Comprehensive logging to file with timestamps
- CSV export option alongside Excel
- Both Excel and CSV export simultaneously
- Enhanced error handling and validation
- Parameter sets for different invocation methods
- MaxRetries parameter for customizable retry attempts
- Detailed audit statistics in summary output
- AuditDate timestamp on all exported records
- Run-Audit.ps1 launcher script for simplified execution

### Improved
- Better module installation with version checking
- More robust connection handling with retry logic
- Enhanced permission detection logic
- Improved large list handling with batch processing
- Better progress feedback during long-running operations
- More detailed console output with color coding
- Optimized memory usage for large audits
- Enhanced documentation with comprehensive README

### Changed
- Refactored main audit function into smaller, reusable components
- Improved error messages with more context
- Better handling of disconnections and timeouts
- Enhanced group membership expansion logic
- Standardized output formatting

### Fixed
- Issue with very large lists causing timeouts
- Memory leaks during long-running audits
- Incorrect handling of nested groups
- Edge cases with special characters in titles
- Connection persistence issues across multiple sites

## [1.0.0] - 2024-XX-XX

### Added
- Initial release
- Basic SharePoint permissions auditing
- Site, list, library, folder, and item level permissions
- Excel export functionality
- SharePoint group membership expansion
- Interactive authentication support
- PnP.PowerShell integration

### Features
- Multi-site auditing capability
- Unique permissions detection
- Direct vs. group-based permission tracking
- Excel export with formatted tables
- Summary statistics

---

## Future Enhancements

### Planned for v2.1
- [ ] Differential auditing (compare two audit results)
- [ ] Permission recommendations and cleanup suggestions
- [ ] HTML report generation
- [ ] Email notification on completion
- [ ] Integration with Azure AD for enhanced user details
- [ ] Support for hub sites and site collections
- [ ] Tenant-wide auditing capability

### Under Consideration
- [ ] PowerShell Gallery publication
- [ ] GUI interface option
- [ ] Scheduled task templates
- [ ] API for programmatic access
- [ ] Integration with compliance tools
- [ ] Custom permission level detection
- [ ] Sharing link auditing
- [ ] External user tracking
