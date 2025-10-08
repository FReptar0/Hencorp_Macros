# Hencorp Macros - Sage 300 ERP Integration Suite

A comprehensive collection of Visual Basic 6.0 automation tools designed for seamless integration with Sage 300 ERP system. This suite provides automated data processing capabilities for Accounts Payable (AP), Accounts Receivable (AR), and General Ledger (GL) modules, along with advanced reporting functionality.

## System Overview

This repository contains six main applications that interface with Sage 300 through COM API integration:

### Core Modules

**MacroAP** - Accounts Payable Processing

- Automated payment entry creation and batch processing
- Bank validation and payment code management
- Error logging and transaction status tracking
- Batch posting capabilities

**MacroAR** - Accounts Receivable Processing  

- Cash receipt processing and bank deposit management
- Customer payment allocation and reference tracking
- Multi-currency transaction support
- Automated posting workflow

**MacroGL** - General Ledger Integration

- Journal entry automation and batch management
- Account validation and fiscal period handling
- Transaction reference tracking
- Automated posting to GL

### Reporting Modules

**Reporte AP** - Accounts Payable Reports
**Reporte AR** - Accounts Receivable Reports  
**Reporte GL** - General Ledger Reports

## Technical Architecture

### Technology Stack

- **Language**: Visual Basic 6.0
- **Database**: SQL Server with ADODB connectivity
- **ERP Integration**: Sage 300 COM API (AccpacCOMAPI)
- **Configuration**: INI file-based settings management

### Key Dependencies

- Sage 300 RUNTIME environment
- ACCPAC COM API Object 1.0
- ACCPAC Session Manager 1.0
- ACCPAC Signon Manager 3.0
- Microsoft ActiveX Data Objects 2.8

## Directory Structure

```tree
Hencorp_Macros/
├── MacroAP/           # Accounts Payable module
├── MacroAR/           # Accounts Receivable module  
├── MacroGL/           # General Ledger module
├── Reporte AP/        # AP reporting application
├── Reporte AR/        # AR reporting application
└── Reporte GL/        # GL reporting application
```

### Standard Module Components

Each module contains:

- **Main executable** (.exe) - Compiled application
- **VB project files** (.vbp, .vbw) - Visual Basic project configuration
- **Form files** (.frm, .frx) - User interface components
- **Module files** (.bas) - Core business logic
- **Config.ini** - Application configuration settings
- **Icon files** (.ico) - Application icons

## Core Functionality

### Data Processing Workflow

1. **Configuration Loading** - Reads database and Sage connection parameters
2. **Database Connection** - Establishes SQL Server connectivity
3. **Sage Session Management** - Authenticates and opens Sage 300 session
4. **Data Retrieval** - Queries staging tables for pending transactions
5. **Batch Processing** - Creates and processes batches in Sage 300
6. **Error Handling** - Logs errors and updates transaction status
7. **Auto-Posting** - Posts completed batches to live data

### Key Features

- **Batch Management**: Automatic creation and processing of transaction batches
- **Error Recovery**: Comprehensive error logging with detailed audit trails
- **Status Tracking**: Real-time transaction status monitoring
- **Data Validation**: Built-in validation for accounts, banks, and references
- **Multi-Company Support**: Handles multiple company databases
- **Automated Posting**: Optional automatic posting of completed batches

## Configuration

### Database Connection Settings

```ini
[settings]
server=DATABASE_SERVER
user=DB_USERNAME  
password=DB_PASSWORD
dbinformacion=DATABASE_NAME
```

### Sage 300 Authentication

```ini
userSage=SAGE_USERNAME
PassSage=SAGE_PASSWORD
```

### Processing Controls

```ini
AsientaAP=SI    # Enable AP auto-posting
AsientaAR=SI    # Enable AR auto-posting  
AsientaGL=SI    # Enable GL auto-posting
log=LOG_PATH    # Error log directory
```

## Database Schema Requirements

### Staging Tables

- **AP_PA** - AP payment headers with status tracking
- **AP_MP** - AP payment details with account distributions
- **AR_RA** - AR receipt headers with deposit information
- **AR_MR** - AR receipt details with account allocations
- **GL_JH** - GL journal headers with batch information
- **GL_JD** - GL journal details with transaction lines

### Status Management

Each staging table includes:

- **ESTADO** - Processing status (null/Completo/Error)
- **RESULTADO** - Error message details
- **LOTE/ASIENTO** - Sage batch and entry numbers
- **FECHA/HORA/USUARIO** - Audit trail information

## Security Considerations

:bangbang: **Important Security Notes**

- Configuration files contain database credentials in plain text
- Sage 300 passwords are stored in configuration files
- Applications run with elevated database permissions
- Error logs may contain sensitive transaction data
- Network traffic includes unencrypted authentication

### Recommended Security Measures

1. Restrict file system access to configuration directories
2. Use dedicated service accounts with minimal permissions
3. Implement secure credential management practices
4. Monitor and secure error log directories
5. Use encrypted network connections where possible

## Operational Requirements

### System Prerequisites

- Windows environment with VB6 runtime
- Sage 300 client installation and licensing
- SQL Server connectivity and appropriate permissions
- Network access to Sage 300 server infrastructure

### Performance Considerations

- Batch size limitations based on system resources
- Database connection pooling for high-volume processing
- Sage 300 session management and timeout handling
- Error recovery and retry mechanisms

## Maintenance and Support

### Monitoring

- Regular review of error logs for system issues
- Database connection health monitoring
- Sage 300 session management oversight
- Batch processing performance tracking

### Troubleshooting

- Check Config.ini for correct connection parameters
- Verify Sage 300 user permissions and module access
- Review SQL Server connectivity and database permissions
- Examine error logs for detailed failure information

---

**Note**: This system is designed for enterprise ERP integration and requires appropriate technical expertise for deployment, configuration, and maintenance in production environments.
**Disclaimer**: Use this software at your own risk. Ensure compliance with your organization's IT policies and Sage 300 licensing agreements.
**License**: MIT License
**Author**: Tersoft - [www.tersoft.mx](http://www.tersoft.mx)
**Version**: 1.0.0 - October 2025

---

