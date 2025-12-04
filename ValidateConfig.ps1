# ValidateConfig.ps1
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  验证 SAP B1 连接配置" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$configPath = ".\config.json"

if (-not (Test-Path $configPath)) {
    Write-Host "配置文件不存在: $configPath" -ForegroundColor Red
    exit 1
}

# 读取配置
$config = Get-Content $configPath | ConvertFrom-Json

Write-Host "配置文件内容:" -ForegroundColor Yellow
Write-Host "  ServerName: $($config.ServerName)"
Write-Host "  DatabaseType: $($config.DatabaseType)"
Write-Host "  DatabaseUserName: $($config.DatabaseUserName)"
Write-Host "  Company: $($config.Company)"
Write-Host "  B1Username: $($config.B1Username)"
Write-Host "  SLDServer: $($config.SLDServer)"
Write-Host ""

# 测试 SQL Server 连接
Write-Host "测试 SQL Server 连接..." -ForegroundColor Yellow

$sqlServer = $config.ServerName -replace ";.*", ""
$dbUser = $config.DatabaseUserName
$dbPass = $config.DatabasePassword
$dbName = $config.Company

$connectionString = "Server=$sqlServer;Database=$dbName;User Id=$dbUser;Password=$dbPass;TrustServerCertificate=True;"

Write-Host "连接字符串: Server=$sqlServer;Database=$dbName;User Id=$dbUser;Password=****" -ForegroundColor Gray

try {
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $sqlConnection.ConnectionString = $connectionString
    $sqlConnection.Open()
    
    Write-Host "  ✓ SQL Server 连接成功" -ForegroundColor Green
    
    # 查询数据库版本
    $sqlCmd = $sqlConnection.CreateCommand()
    $sqlCmd.CommandText = "SELECT @@VERSION"
    $version = $sqlCmd.ExecuteScalar()
    Write-Host "  SQL Server 版本: $($version.Split("`n")[0])" -ForegroundColor Gray
    
    # 检查公司数据库
    $sqlCmd.CommandText = "SELECT DB_NAME()"
    $currentDb = $sqlCmd.ExecuteScalar()
    Write-Host "  当前数据库: $currentDb" -ForegroundColor Gray
    
    # 检查用户表是否存在
    $sqlCmd.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'OUSR'"
    $tableExists = $sqlCmd.ExecuteScalar()
    
    if ($tableExists -gt 0) {
        Write-Host "  ✓ SAP B1 用户表存在" -ForegroundColor Green
        
        # 查询用户
        $b1User = $config.B1Username
        $sqlCmd.CommandText = "SELECT USER_CODE, U_NAME, LOCKED FROM OUSR WHERE USER_CODE = '$b1User'"
        $reader = $sqlCmd.ExecuteReader()
        
        if ($reader.Read()) {
            $userName = $reader["U_NAME"]
            $locked = $reader["LOCKED"]
            $reader.Close()
            
            Write-Host "  找到用户: $userName (USER_CODE: $b1User)" -ForegroundColor Green
            
            if ($locked -eq "Y") {
                Write-Host "  ✗ 警告: 用户账户已被锁定！" -ForegroundColor Red
            } else {
                Write-Host "  ✓ 用户账户状态正常" -ForegroundColor Green
            }
        } else {
            $reader.Close()
            Write-Host "  ✗ 未找到用户: $b1User" -ForegroundColor Red
            Write-Host "    请检查用户名是否正确" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  ✗ 这不是一个有效的 SAP B1 数据库" -ForegroundColor Red
    }
    
    $sqlConnection.Close()
}
catch {
    Write-Host "  ✗ SQL Server 连接失败" -ForegroundColor Red
    Write-Host "  错误: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "验证完成" -ForegroundColor Cyan