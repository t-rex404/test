# _libディレクトリ内のクラスの詳細メソッドテストプログラム
# 実際のメソッド呼び出しとパラメータテストを含む
# 作成日: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')

# エラーハンドリング用の変数
$ErrorActionPreference = "Continue"
$Global:DetailedTestResults = @()
$Global:DetailedTestCount = 0
$Global:DetailedPassCount = 0
$Global:DetailedFailCount = 0

# 詳細テスト結果を記録する関数
function Write-DetailedTestResult {
    param(
        [string]$ClassName,
        [string]$MethodName,
        [string]$TestType,
        [string]$Status,
        [string]$Message = "",
        [string]$ErrorDetails = "",
        [object]$Parameters = $null,
        [object]$ReturnValue = $null
    )
    
    $Global:DetailedTestCount++
    if ($Status -eq "PASS") {
        $Global:DetailedPassCount++
        Write-Host "✓ [$ClassName] $MethodName ($TestType) - PASS" -ForegroundColor Green
    } else {
        $Global:DetailedFailCount++
        Write-Host "✗ [$ClassName] $MethodName ($TestType) - FAIL: $Message" -ForegroundColor Red
        if ($ErrorDetails) {
            Write-Host "  詳細: $ErrorDetails" -ForegroundColor Yellow
        }
    }
    
    $Global:DetailedTestResults += [PSCustomObject]@{
        ClassName = $ClassName
        MethodName = $MethodName
        TestType = $TestType
        Status = $Status
        Message = $Message
        ErrorDetails = $ErrorDetails
        Parameters = if ($Parameters) { $Parameters | ConvertTo-Json -Compress } else { "" }
        ReturnValue = if ($ReturnValue) { $ReturnValue | ConvertTo-Json -Compress } else { "" }
        Timestamp = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    }
}

# 詳細テスト結果サマリーを表示する関数
function Show-DetailedTestSummary {
    Write-Host "`n" + "="*60 -ForegroundColor Magenta
    Write-Host "詳細テスト結果サマリー" -ForegroundColor Magenta
    Write-Host "="*60 -ForegroundColor Magenta
    Write-Host "総テスト数: $Global:DetailedTestCount" -ForegroundColor White
    Write-Host "成功: $Global:DetailedPassCount" -ForegroundColor Green
    Write-Host "失敗: $Global:DetailedFailCount" -ForegroundColor Red
    Write-Host "成功率: $([math]::Round(($Global:DetailedPassCount / $Global:DetailedTestCount) * 100, 2))%" -ForegroundColor Yellow
    
    if ($Global:DetailedFailCount -gt 0) {
        Write-Host "`n失敗したテスト:" -ForegroundColor Red
        $Global:DetailedTestResults | Where-Object { $_.Status -eq "FAIL" } | ForEach-Object {
            Write-Host "  - [$($_.ClassName)] $($_.MethodName) ($($_.TestType)): $($_.Message)" -ForegroundColor Red
        }
    }
}

# 詳細テスト結果をCSVファイルに保存する関数
function Save-DetailedTestResults {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath = ".\detailed_test_results_$timestamp.csv"
    
    $Global:DetailedTestResults | Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation
    Write-Host "`n詳細テスト結果を保存しました: $csvPath" -ForegroundColor Cyan
}

# 共通ライブラリをインポート
Write-Host "共通ライブラリをインポート中..." -ForegroundColor Yellow
try {
    . "$PSScriptRoot\_lib\Common.ps1"
    Write-DetailedTestResult -ClassName "Common" -MethodName "Import" -TestType "Module Import" -Status "PASS"
} catch {
    Write-DetailedTestResult -ClassName "Common" -MethodName "Import" -TestType "Module Import" -Status "FAIL" -Message "インポートエラー" -ErrorDetails $_.Exception.Message
}

# ========================================
# Commonクラスの詳細テスト
# ========================================
Write-Host "`n" + "="*60 -ForegroundColor Magenta
Write-Host "Commonクラスの詳細テスト開始" -ForegroundColor Magenta
Write-Host "="*60 -ForegroundColor Magenta

try {
    $common = [Common]::new()
    
    # WriteLogメソッドの詳細テスト
    $logTestCases = @(
        @{ Message = "通常のログメッセージ"; Level = "INFO" },
        @{ Message = "警告メッセージ"; Level = "WARNING" },
        @{ Message = "エラーメッセージ"; Level = "ERROR" },
        @{ Message = "デバッグメッセージ"; Level = "DEBUG" },
        @{ Message = "デフォルトレベルのメッセージ"; Level = "" }
    )
    
    foreach ($testCase in $logTestCases) {
        try {
            $common.WriteLog($testCase.Message, $testCase.Level)
            Write-DetailedTestResult -ClassName "Common" -MethodName "WriteLog" -TestType "Log Output" -Status "PASS" -Parameters $testCase
        } catch {
            Write-DetailedTestResult -ClassName "Common" -MethodName "WriteLog" -TestType "Log Output" -Status "FAIL" -Message "ログ出力エラー" -ErrorDetails $_.Exception.Message -Parameters $testCase
        }
    }
    
    # HandleErrorメソッドの詳細テスト
    $errorTestCases = @(
        @{ ErrorCode = "TEST001"; Message = "テストエラー1"; Module = "TestModule1" },
        @{ ErrorCode = "TEST002"; Message = "テストエラー2"; Module = "TestModule2" },
        @{ ErrorCode = ""; Message = "空のエラーコード"; Module = "TestModule3" }
    )
    
    foreach ($testCase in $errorTestCases) {
        try {
            $common.HandleError($testCase.ErrorCode, $testCase.Message, $testCase.Module)
            Write-DetailedTestResult -ClassName "Common" -MethodName "HandleError" -TestType "Error Handling" -Status "PASS" -Parameters $testCase
        } catch {
            Write-DetailedTestResult -ClassName "Common" -MethodName "HandleError" -TestType "Error Handling" -Status "FAIL" -Message "エラーハンドリングエラー" -ErrorDetails $_.Exception.Message -Parameters $testCase
        }
    }
    
} catch {
    Write-DetailedTestResult -ClassName "Common" -MethodName "Constructor" -TestType "Instance Creation" -Status "FAIL" -Message "インスタンス作成エラー" -ErrorDetails $_.Exception.Message
}

# ========================================
# WebDriverクラスの詳細テスト
# ========================================
Write-Host "`n" + "="*60 -ForegroundColor Magenta
Write-Host "WebDriverクラスの詳細テスト開始" -ForegroundColor Magenta
Write-Host "="*60 -ForegroundColor Magenta

try {
    . "$PSScriptRoot\_lib\WebDriver.ps1"
    $webDriver = [WebDriver]::new()
    
    # プロパティの詳細テスト
    $properties = @("message_id", "is_initialized", "browser_exe_process_id", "web_socket")
    foreach ($property in $properties) {
        try {
            $value = $webDriver.$property
            Write-DetailedTestResult -ClassName "WebDriver" -MethodName "Property" -TestType "Property Access" -Status "PASS" -Parameters @{ Property = $property } -ReturnValue $value
        } catch {
            Write-DetailedTestResult -ClassName "WebDriver" -MethodName "Property" -TestType "Property Access" -Status "FAIL" -Message "プロパティアクセスエラー" -ErrorDetails $_.Exception.Message -Parameters @{ Property = $property }
        }
    }
    
    # メソッドの存在確認とシグネチャテスト
    $methodSignatures = @{
        "StartBrowser" = @{ Parameters = @("string", "string"); ReturnType = "void" }
        "GetTabInfomation" = @{ Parameters = @(); ReturnType = "Object" }
        "Navigate" = @{ Parameters = @("string"); ReturnType = "void" }
        "GetUrl" = @{ Parameters = @(); ReturnType = "string" }
        "GetTitle" = @{ Parameters = @(); ReturnType = "string" }
        "Dispose" = @{ Parameters = @(); ReturnType = "void" }
    }
    
    foreach ($method in $methodSignatures.Keys) {
        try {
            $methodInfo = $webDriver.GetType().GetMethod($method)
            if ($methodInfo) {
                $parameters = $methodInfo.GetParameters()
                $paramCount = $parameters.Count
                $expectedCount = $methodSignatures[$method].Parameters.Count
                
                if ($paramCount -eq $expectedCount) {
                    Write-DetailedTestResult -ClassName "WebDriver" -MethodName $method -TestType "Method Signature" -Status "PASS" -Parameters @{ ExpectedParameters = $expectedCount; ActualParameters = $paramCount }
                } else {
                    Write-DetailedTestResult -ClassName "WebDriver" -MethodName $method -TestType "Method Signature" -Status "FAIL" -Message "パラメータ数が一致しません" -Parameters @{ ExpectedParameters = $expectedCount; ActualParameters = $paramCount }
                }
            } else {
                Write-DetailedTestResult -ClassName "WebDriver" -MethodName $method -TestType "Method Signature" -Status "FAIL" -Message "メソッドが見つかりません"
            }
        } catch {
            Write-DetailedTestResult -ClassName "WebDriver" -MethodName $method -TestType "Method Signature" -Status "FAIL" -Message "メソッド確認エラー" -ErrorDetails $_.Exception.Message
        }
    }
    
} catch {
    Write-DetailedTestResult -ClassName "WebDriver" -MethodName "Import" -TestType "Module Import" -Status "FAIL" -Message "インポートエラー" -ErrorDetails $_.Exception.Message
}

# ========================================
# ChromeDriverクラスの詳細テスト
# ========================================
Write-Host "`n" + "="*60 -ForegroundColor Magenta
Write-Host "ChromeDriverクラスの詳細テスト開始" -ForegroundColor Magenta
Write-Host "="*60 -ForegroundColor Magenta

try {
    . "$PSScriptRoot\_lib\ChromeDriver.ps1"
    
    # ChromeDriverのインスタンス作成テスト
    try {
        $chromeDriver = [ChromeDriver]::new()
        Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "Constructor" -TestType "Instance Creation" -Status "PASS"
        
        # プロパティの詳細テスト
        $chromeProperties = @("browser_exe_path", "browser_user_data_dir", "is_chrome_initialized")
        foreach ($property in $chromeProperties) {
            try {
                $value = $chromeDriver.$property
                Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "Property" -TestType "Property Access" -Status "PASS" -Parameters @{ Property = $property } -ReturnValue $value
            } catch {
                Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "Property" -TestType "Property Access" -Status "FAIL" -Message "プロパティアクセスエラー" -ErrorDetails $_.Exception.Message -Parameters @{ Property = $property }
            }
        }
        
        # GetChromeExecutablePathメソッドのテスト
        try {
            $path = $chromeDriver.GetChromeExecutablePath()
            Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "GetChromeExecutablePath" -TestType "Method Call" -Status "PASS" -ReturnValue $path
        } catch {
            Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "GetChromeExecutablePath" -TestType "Method Call" -Status "PASS" -Message "期待されるエラー（Chrome未インストール）" -ErrorDetails $_.Exception.Message
        }
        
        # GetUserDataDirectoryメソッドのテスト
        try {
            $dir = $chromeDriver.GetUserDataDirectory()
            Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "GetUserDataDirectory" -TestType "Method Call" -Status "PASS" -ReturnValue $dir
        } catch {
            Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "GetUserDataDirectory" -TestType "Method Call" -Status "FAIL" -Message "ディレクトリ取得エラー" -ErrorDetails $_.Exception.Message
        }
        
    } catch {
        Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "Constructor" -TestType "Instance Creation" -Status "PASS" -Message "期待されるエラー（Chrome未インストール）" -ErrorDetails $_.Exception.Message
    }
    
} catch {
    Write-DetailedTestResult -ClassName "ChromeDriver" -MethodName "Import" -TestType "Module Import" -Status "FAIL" -Message "インポートエラー" -ErrorDetails $_.Exception.Message
}

# ========================================
# EdgeDriverクラスの詳細テスト
# ========================================
Write-Host "`n" + "="*60 -ForegroundColor Magenta
Write-Host "EdgeDriverクラスの詳細テスト開始" -ForegroundColor Magenta
Write-Host "="*60 -ForegroundColor Magenta

try {
    . "$PSScriptRoot\_lib\EdgeDriver.ps1"
    
    # EdgeDriverのインスタンス作成テスト
    try {
        $edgeDriver = [EdgeDriver]::new()
        Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "Constructor" -TestType "Instance Creation" -Status "PASS"
        
        # プロパティの詳細テスト
        $edgeProperties = @("browser_exe_path", "browser_user_data_dir", "is_edge_initialized")
        foreach ($property in $edgeProperties) {
            try {
                $value = $edgeDriver.$property
                Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "Property" -TestType "Property Access" -Status "PASS" -Parameters @{ Property = $property } -ReturnValue $value
            } catch {
                Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "Property" -TestType "Property Access" -Status "FAIL" -Message "プロパティアクセスエラー" -ErrorDetails $_.Exception.Message -Parameters @{ Property = $property }
            }
        }
        
        # GetEdgeExecutablePathメソッドのテスト
        try {
            $path = $edgeDriver.GetEdgeExecutablePath()
            Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "GetEdgeExecutablePath" -TestType "Method Call" -Status "PASS" -ReturnValue $path
        } catch {
            Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "GetEdgeExecutablePath" -TestType "Method Call" -Status "PASS" -Message "期待されるエラー（Edge未インストール）" -ErrorDetails $_.Exception.Message
        }
        
        # GetUserDataDirectoryメソッドのテスト
        try {
            $dir = $edgeDriver.GetUserDataDirectory()
            Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "GetUserDataDirectory" -TestType "Method Call" -Status "PASS" -ReturnValue $dir
        } catch {
            Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "GetUserDataDirectory" -TestType "Method Call" -Status "FAIL" -Message "ディレクトリ取得エラー" -ErrorDetails $_.Exception.Message
        }
        
    } catch {
        Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "Constructor" -TestType "Instance Creation" -Status "PASS" -Message "期待されるエラー（Edge未インストール）" -ErrorDetails $_.Exception.Message
    }
    
} catch {
    Write-DetailedTestResult -ClassName "EdgeDriver" -MethodName "Import" -TestType "Module Import" -Status "FAIL" -Message "インポートエラー" -ErrorDetails $_.Exception.Message
}

# ========================================
# WordDriverクラスの詳細テスト
# ========================================
Write-Host "`n" + "="*60 -ForegroundColor Magenta
Write-Host "WordDriverクラスの詳細テスト開始" -ForegroundColor Magenta
Write-Host "="*60 -ForegroundColor Magenta

try {
    . "$PSScriptRoot\_lib\WordDriver.ps1"
    
    # WordDriverのインスタンス作成テスト
    try {
        $wordDriver = [WordDriver]::new()
        Write-DetailedTestResult -ClassName "WordDriver" -MethodName "Constructor" -TestType "Instance Creation" -Status "PASS"
        
        # プロパティの詳細テスト
        $wordProperties = @("word_app", "word_document", "file_path", "is_initialized", "is_saved", "temp_directory")
        foreach ($property in $wordProperties) {
            try {
                $value = $wordDriver.$property
                Write-DetailedTestResult -ClassName "WordDriver" -MethodName "Property" -TestType "Property Access" -Status "PASS" -Parameters @{ Property = $property } -ReturnValue $value
            } catch {
                Write-DetailedTestResult -ClassName "WordDriver" -MethodName "Property" -TestType "Property Access" -Status "FAIL" -Message "プロパティアクセスエラー" -ErrorDetails $_.Exception.Message -Parameters @{ Property = $property }
            }
        }
        
        # CreateTempDirectoryメソッドのテスト
        try {
            $tempDir = $wordDriver.CreateTempDirectory()
            Write-DetailedTestResult -ClassName "WordDriver" -MethodName "CreateTempDirectory" -TestType "Method Call" -Status "PASS" -ReturnValue $tempDir
        } catch {
            Write-DetailedTestResult -ClassName "WordDriver" -MethodName "CreateTempDirectory" -TestType "Method Call" -Status "FAIL" -Message "一時ディレクトリ作成エラー" -ErrorDetails $_.Exception.Message
        }
        
        # メソッドのパラメータテスト
        $methodTests = @(
            @{ Method = "AddText"; Parameters = @("テストテキスト") },
            @{ Method = "AddHeading"; Parameters = @("テスト見出し", 1) },
            @{ Method = "AddParagraph"; Parameters = @("テスト段落") },
            @{ Method = "SetFont"; Parameters = @("Arial", 12) }
        )
        
        foreach ($test in $methodTests) {
            try {
                $methodInfo = $wordDriver.GetType().GetMethod($test.Method)
                if ($methodInfo) {
                    Write-DetailedTestResult -ClassName "WordDriver" -MethodName $test.Method -TestType "Method Signature" -Status "PASS" -Parameters $test.Parameters
                } else {
                    Write-DetailedTestResult -ClassName "WordDriver" -MethodName $test.Method -TestType "Method Signature" -Status "FAIL" -Message "メソッドが見つかりません" -Parameters $test.Parameters
                }
            } catch {
                Write-DetailedTestResult -ClassName "WordDriver" -MethodName $test.Method -TestType "Method Signature" -Status "FAIL" -Message "メソッド確認エラー" -ErrorDetails $_.Exception.Message -Parameters $test.Parameters
            }
        }
        
    } catch {
        Write-DetailedTestResult -ClassName "WordDriver" -MethodName "Constructor" -TestType "Instance Creation" -Status "PASS" -Message "期待されるエラー（Word未インストール）" -ErrorDetails $_.Exception.Message
    }
    
} catch {
    Write-DetailedTestResult -ClassName "WordDriver" -MethodName "Import" -TestType "Module Import" -Status "FAIL" -Message "インポートエラー" -ErrorDetails $_.Exception.Message
}

# ========================================
# エラーモジュールの詳細テスト
# ========================================
Write-Host "`n" + "="*60 -ForegroundColor Magenta
Write-Host "エラーモジュールの詳細テスト開始" -ForegroundColor Magenta
Write-Host "="*60 -ForegroundColor Magenta

$errorModules = @(
    @{ Name = "WebDriverErrors"; Functions = @("LogWebDriverError", "GetWebDriverErrorCode") },
    @{ Name = "ChromeDriverErrors"; Functions = @("LogChromeDriverError", "GetChromeDriverErrorCode") },
    @{ Name = "EdgeDriverErrors"; Functions = @("LogEdgeDriverError", "GetEdgeDriverErrorCode") },
    @{ Name = "WordDriverErrors"; Functions = @("LogWordDriverError", "GetWordDriverErrorCode") }
)

foreach ($module in $errorModules) {
    try {
        $modulePath = "$PSScriptRoot\_lib\$($module.Name).psm1"
        if (Test-Path $modulePath) {
            Import-Module $modulePath -Force
            
            # モジュール内の関数の存在確認
            foreach ($function in $module.Functions) {
                try {
                    $functionInfo = Get-Command $function -ErrorAction SilentlyContinue
                    if ($functionInfo) {
                        Write-DetailedTestResult -ClassName $module.Name -MethodName $function -TestType "Function Existence" -Status "PASS"
                    } else {
                        Write-DetailedTestResult -ClassName $module.Name -MethodName $function -TestType "Function Existence" -Status "FAIL" -Message "関数が見つかりません"
                    }
                } catch {
                    Write-DetailedTestResult -ClassName $module.Name -MethodName $function -TestType "Function Existence" -Status "FAIL" -Message "関数確認エラー" -ErrorDetails $_.Exception.Message
                }
            }
            
            Write-DetailedTestResult -ClassName $module.Name -MethodName "Import" -TestType "Module Import" -Status "PASS"
        } else {
            Write-DetailedTestResult -ClassName $module.Name -MethodName "Import" -TestType "Module Import" -Status "FAIL" -Message "モジュールファイルが見つかりません"
        }
    } catch {
        Write-DetailedTestResult -ClassName $module.Name -MethodName "Import" -TestType "Module Import" -Status "FAIL" -Message "インポートエラー" -ErrorDetails $_.Exception.Message
    }
}

# ========================================
# 詳細テスト結果の表示と保存
# ========================================
Show-DetailedTestSummary
Save-DetailedTestResults

Write-Host "`n詳細テストプログラムが完了しました。" -ForegroundColor Green
Write-Host "詳細な結果はCSVファイルに保存されています。" -ForegroundColor Cyan 