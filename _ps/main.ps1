$Host.UI.RawUI.WindowTitle = 'TEST'
Add-Type -AssemblyName Microsoft.VisualBasic
write-host "TEST'$PSScriptRoot"
. $PSScriptRoot\_lib\WebDriver.ps1
. $PSScriptRoot\_lib\ChromeDriver.ps1
. $PSScriptRoot\_lib\EdgeDriver.ps1



Function Main()
{
    param
    (
        
    )

    begin
    {
        # 定数宣言
        new-variable -name SUCCESS_CODE -value 0 -Description '正常終了コード' -Option Constant -visibility Private
        new-variable -name ERROR_CODE -value 255 -Description '異常終了コード' -Option Constant -visibility Private

        # 変数宣言
        new-variable -name result -value $ERROR_CODE -Description '結果' -Option Private -visibility Private
    }

    process
    {
        try
        {
            $url_list = 'https://www.fe-siken.com/index_te.html', 'https://www.fe-siken.com/index_ma.html', 'https://www.fe-siken.com/index_st.html', 'https://www.fe-siken.com/index_pm.html'
            $major_list = 'テクノロジ系', 'マネジメント系', 'ストラテジ系', '午後問題の分野別まとめ'
            # EddgeDriver インスタンス生成
            #$driver = [EdgeDriver]::new
            $driver = New-Object -TypeName 'EdgeDriver'

            for ($x = 0; $x -lt $url_list.count; $x++)
            {
                #write-host 'Navigate'
                $driver.Navigate($url_list[$x])
                #Start-Sleep -Seconds 1

                $count = 0
                for ($i = 1; $i -lt 10000; $i++)
                {
                    if ($x -eq 3)
                    {
                        $xpath_0 = "/html/body/div[1]/div/main/div[2]/div[1]/div[1]/dl[$i]/dt"
                    }
                    else
                    {
                        $xpath_0 = "/html/body/div[1]/div/main/div[2]/div[2]/div[1]/dl[$i]/dt"
                    }

                    if ($driver.IsExistsElementByXPath($xpath_0) -eq $false)
                    {
                        break
                    }
                    $element = $driver.FindElementByXpath($xpath_0)
                    if ($element)
                    {
                        $medium_text = $driver.GetElementText($element.nodeId)
                        #write-host $medium_text
                    
                        for ($j = 1; $j -lt 10000; $j++)
                        {
                            if ($x -eq 3)
                            {
                                $xpath_1 = "/html/body/div[1]/div/main/div[2]/div[1]/div[1]/dl[$i]/dd/a[$j]"
                            }
                            else
                            {
                                $xpath_1 = "/html/body/div[1]/div/main/div[2]/div[2]/div[1]/dl[$i]/dd/a[$j]"
                            }

                            if ($driver.IsExistsElementByXPath($xpath_1) -eq $false)
                            {
                                break
                            }
                            $element = $driver.FindElementByXpath($xpath_1)
                            if ($element)
                            {
                                $minor_text = $driver.GetElementText($element.nodeId)
                                #write-host "$medium_text; $minor_text"
                                #$driver.ClickElement($element.nodeId)
                            }
                            $count++
                            for ($k = 2; $k -lt 10000; $k++)
                            {
                                if ($x -eq 3)
                                {
                                    $xpath_2 = "/html/body/div[1]/div/main/div[2]/div[1]/div[2]/table[$count]/tbody/tr[$k]/td[2]/a"
                                }
                                else
                                {
                                    $xpath_2 = "/html/body/div[1]/div/main/div[2]/div[2]/div[2]/table[$count]/tbody/tr[$k]/td[2]/a"
                                }
                                if ($driver.IsExistsElementByXPath($xpath_2) -eq $false)
                                {
                                    break
                                }
                                $element = $driver.FindElementByXpath($xpath_2)
                                if ($element.nodeId)
                                {
                                    $minor2_text = $driver.GetElementText($element.nodeId)
                                    write-host "$($major_list[$x]),$medium_text,$minor_text,$minor2_text"
                                    $question_link = $driver.GetHrefFromAnchor($element.nodeId)
                                    $driver.Navigate($question_link)
                                    #$driver.ClickElement($element.nodeId)

                                    <#
                                    if ($x -eq 3)
                                    {
                                        $xpath_3 = "/html/body/div[1]/div/main/div[3]/div[1]"
                                    }
                                    else
                                    {
                                        $xpath_3 = "/html/body/div[1]/div/main/div[2]/div[1]"
                                    }
                                    if ($driver.IsExistsElementByXPath($xpath_3) -eq $false)
                                    {
                                        $question_text = ''
                                        $file_path = '.\kihonjoho.csv'
                                        Out-File -Append -FilePath $file_path -InputObject "$($major_list[$x]),$medium_text,$minor_text,$minor2_text,$question_text"
                                        $driver.Navigate($url_list[$x])
                                    }
                                    else
                                    {
                                        $element = $driver.FindElementByXpath($xpath_3)
                                        if ($element)
                                        {
                                            $question_text = $driver.GetElementText($element.nodeId)
                                            $file_path = '.\kihonjoho.csv'
                                            Out-File -Append -FilePath $file_path -InputObject "$($major_list[$x]),$medium_text,$minor_text,$minor2_text,$question_text" -encoding utf8
                                            $driver.Navigate($url_list[$x])
                                        }
                                    }
                                    #>
                                    if ($x -eq 3)
                                    {
                                        $question_xpath_3    = "/html/body/div[1]/div/main/div[3]/div[1]"
                                        $choices_xpath_3     = "/html/body/div[1]/div/main/div[3]/div[2]"
                                        $answer_xpath_3      = "/html/body/div[1]/div/main/div[3]/div[3]"
                                        $explanation_xpath_3 = "/html/body/div[1]/div/main/div[3]/div[3]"
                                    }
                                    else
                                    {
                                        $question_xpath_3    = "/html/body/div[1]/div/main/div[2]/div[1]"
                                        $choices_xpath_3     = "/html/body/div[1]/div/main/div[2]/div[3]/ul"
                                        $answer_xpath_3      = "/html/body/div[1]/div/main/div[2]/div[5]/span"
                                        $explanation_xpath_3 = "/html/body/div[1]/div/main/div[2]/div[6]"
                                    }
                                    if ($driver.IsExistsElementByXPath($question_xpath_3) -eq $false)
                                    {
                                        #return 0
                                    }
                                    else
                                    {
                                        $question_png_file = "question_$($x + 1)_$i`_$j`_$($k - 1).png"
                                        $choices_png_file  = "choices_$($x + 1)_$i`_$j`_$($k - 1).png"
                                        $explanation_png_file = "explanation_$($x + 1)_$i`_$j`_$($k - 1).png"

                                        $question_png_path    = ".\_png\question\$question_png_file"
                                        $choices_png_path     = ".\_png\choices\$choices_png_file"
                                        $explanation_png_path = ".\_png\explanation\$explanation_png_file"

                                        $element = $driver.FindElementByXpath($question_xpath_3)
                                        $driver.GetScreenshotObjectId($element.nodeId, $question_png_path)

                                        $element = $driver.FindElementByXpath($choices_xpath_3)
                                        $driver.GetScreenshotObjectId($element.nodeId, $choices_png_path)

                                        $element = $driver.FindElementByXpath($explanation_xpath_3)
                                        $driver.GetScreenshotObjectId($element.nodeId, $explanation_png_path)

                                        $element = $driver.FindElementByXpath($answer_xpath_3)
                                        $answer_text = $driver.GetElementText($element.nodeId)
                                    }

                                    $file_path = '.\kihonjoho.csv'
                                    Out-File -Append -FilePath $file_path -InputObject "$($major_list[$x]),$medium_text,$minor_text,$minor2_text,$question_png_file,$choices_png_file,$answer_text,$explanation_png_file" -encoding utf8
                                    $driver.Navigate($url_list[$x])
                                }
                            }
                        }
                    }
                }
            }



            <# テストのため一旦コメントアウト
            for ($x = 0; $x -lt $url_list.count; $x++)
            {
                write-host 'Navigate'
                $driver.Navigate($url_list[$x])
                #Start-Sleep -Seconds 1

                $count = 0
                for ($i = 1; $i -lt 10000; $i++)
                {
                    if ($driver.IsExistsElementByXPath("/html/body/div[1]/div/main/div[2]/div[1]/div[1]/dl[$i]/dt") -eq $false)
                    {
                        break
                    }
                    $element = $driver.FindElementByXpath("/html/body/div[1]/div/main/div[2]/div[1]/div[1]/dl[$i]/dt")
                    if ($element)
                    {
                        $medium_text = $driver.GetElementText($element)
                        write-host $medium_text
                    
                        for ($j = 1; $j -lt 10000; $j++)
                        {
                            if ($driver.IsExistsElementByXPath("/html/body/div[1]/div/main/div[2]/div[1]/div[1]/dl[$i]/dd/a[$j]") -eq $false)
                            {
                                break
                            }
                            $element = $driver.FindElementByXpath("/html/body/div[1]/div/main/div[2]/div[1]/div[1]/dl[$i]/dd/a[$j]")
                            if ($element)
                            {
                                $minor_text = $driver.GetElementText($element)
                                write-host "$medium_text; $minor_text"
                            }
                            $count++
                            for ($k = 2; $k -lt 10000; $k++)
                            {
                                
                                if ($driver.IsExistsElementByXPath("/html/body/div[1]/div/main/div[2]/div[1]/div[2]/table[$count]/tbody/tr[$k]/td[2]/a") -eq $false)
                                {
                                    break
                                }
                                $element = $driver.FindElementByXpath("/html/body/div[1]/div/main/div[2]/div[1]/div[2]/table[$count]/tbody/tr[$k]/td[2]/a")
                                if ($element)
                                {
                                    $minor2_text = $driver.GetElementText($element)
                                    write-host "$($major_list[$x]),$medium_text,$minor_text,$minor2_text"
                                    $file_path = '.\kihonjoho.csv'
                                    Out-File -Append -FilePath $file_path -InputObject "$($major_list[$x]),$medium_text,$minor_text,$minor2_text"
                                }
                            }
                        }
                    }
                }
            }
            #>




            write-host 'CloseWindow'
            $driver.CloseWindow()
            write-host 'Close'
            $result = $SUCCESS_CODE
        }
        catch
        {
            #Do this if a terminating exception happens
            write-host 'catch'
            write-host $_.Exception.Message
        }
        finally
        {
            #Do this no matter what
            write-host 'finally'
        }
        return $result
    }
    end
    {
        # 終了処理
        #remove-variable -name 'EDGE_EXE_REG_KEY' -ErrorAction SilentlyContinue
        #remove-variable -name 'result'           -ErrorAction SilentlyContinue
        #remove-variable -name 'edge_exe_path'    -ErrorAction SilentlyContinue
        #remove-variable -name 'process'          -ErrorAction SilentlyContinue    
        #remove-variable -name 'process_id'       -ErrorAction SilentlyContinue
    }
}

$result = Main
switch ($result)
{
    0
    {
        write-host '正常終了'
        exit 0
    }
    1
    {
        write-host '異常終了'
        exit 1
    }
    default
    {
        write-host '異常終了'
        exit 1
    }
}


# 管理者権限が必要ない
# クリックを受け付けなくなって困ったらCTRL+ALT＋DEL
# 中身のメモ帳を動かすスクリプトは以下から
# https://www.tekizai.net/entry/2021/10/02/063000
# フックロジックは以下から
# https://www.ipentec.com/document/csharp-get-mouse-pointer-screen-position-using-global-hook
$cscode = @"
using System;
using System.Runtime.InteropServices;

public class WinHelper 
{
  [StructLayout(LayoutKind.Sequential)]
  struct POINT
  {
    public int x;
    public int y;
  }

  [StructLayout(LayoutKind.Sequential)]
  struct MOUSEHOOKSTRUCT
  {
    public POINT pt;
    public IntPtr hwnd;
    public int wHitTestCode;
    public int dwExtraInfo;
  }

  [StructLayout(LayoutKind.Sequential)]
  struct MOUSEHOOKSTRUCTEX
  {
    public MOUSEHOOKSTRUCT mouseHookStruct;
    public int MouseData;
  }

  [StructLayout(LayoutKind.Sequential)]
  struct MSLLHOOKSTRUCT
  {
    public POINT pt;
    public uint mouseData;
    public uint flags;
    public uint time;
    public IntPtr dwExtraInfo;
  }

  const int HC_ACTION = 0;
  private delegate IntPtr MouseHookCallback(int nCode, uint msg, ref MSLLHOOKSTRUCT msllhookstruct);

  [DllImport("user32.dll")]
  static extern IntPtr SetWindowsHookEx(int idHook, MouseHookCallback lpfn, IntPtr hMod, IntPtr dwThreadId);


  [DllImport("user32.dll")]
  static extern bool UnhookWindowsHookEx(IntPtr hHook);

  [DllImport("user32.dll")]
  static extern IntPtr CallNextHookEx(IntPtr hHook, int nCode, uint wParam, ref MSLLHOOKSTRUCT msllhookstruct);

  private static IntPtr MyHookProc(int nCode, uint wParam, ref MSLLHOOKSTRUCT lParam)
  {
    if (nCode == HC_ACTION) {
      // 0以外を返すことでイベントを握りつぶす
      return (IntPtr)1;
    }
    return CallNextHookEx(hHook, nCode, wParam, ref lParam);
  }
  static IntPtr hHook = IntPtr.Zero;
  static MouseHookCallback proc;
  const int WH_MOUSE_LL = 14;
  public static IntPtr Block()
  {
    proc = MyHookProc;
    hHook = SetWindowsHookEx(WH_MOUSE_LL,proc, IntPtr.Zero, IntPtr.Zero);
    return hHook;
    // return IntPtr.Zero;
  }
  public static bool UnBlock() {
    if (hHook == IntPtr.Zero) return false;
    return UnhookWindowsHookEx(hHook);
  }
}
"@
Add-Type -TypeDefinition $cscode
Write-Output $Win32
# 操作禁止
[WinHelper]::Block()

Write-Output "操作できなくなる"
#ウィンドウの新しい幅 (ピクセル単位)
$width=300
#ウィンドウの新しい高さ (ピクセル単位)。
$height=300
#コントロールの左側の絶対画面座標。
$x=0
#コントロールの上部の絶対画面座標。
$y=0
Add-Type -AssemblyName UIAutomationClient
$notepad=Start-Process notepad -PassThru
Start-Sleep 1
$hwnd=$notepad.MainWindowHandle
#ハンドルからウィンドウを取得する
$window=[System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
#ウィンドウサイズの状態を把握するためにWindowPatternを使う
$windowPattern=$window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
#ウィンドウサイズを変更する準備としてサイズを通常に変更する
$windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
#ウィンドウサイズを変更するためのﾊﾟﾀｰﾝ。
$transformPattern=$window.GetCurrentPattern([System.Windows.Automation.TransformPattern]::Pattern)
#Maximamだと移動もｻｲｽﾞ変更もできないので適当なサイズに変更する。
$transformPattern.Resize($width,$height)
#暴れる間隔
$sleep=0.5
#暴れる秒数
$duration=5
$startTime=Get-Date
while($true)
{
   Start-Sleep $sleep
   $transformPattern.Move($x,$y)
   $x=Get-Random -Maximum 1000 -Minimum 0
   $y=Get-Random -Maximum 1000 -Minimum 0
   $currentTime=Get-Date
   $durationMinute=($currentTime-$startTime).Seconds
   if(($durationMinute -ge $duration) -or $notepad.HasExited){
      Stop-Process $notepad
      break
   }
}

# 操作復活
[WinHelper]::UnBlock()
Write-Output "操作できる"
