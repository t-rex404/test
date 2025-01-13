$Host.UI.RawUI.WindowTitle = 'TEST'
Add-Type -AssemblyName Microsoft.VisualBasic

Function Main()
{
    param
    (
        
    )

    begin
    {
        # 定数宣言
        new-variable -name EDGE_EXE_REG_KEY -value 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\' -Description 'EDGE_EXE_REG_KEY' -Option Constant -visibility Private

        # 変数宣言
        new-variable -name result        -value 255   -Description '結果' -Option Private -visibility Private
        new-variable -name id            -value 0     -Description 'ID' -Option Private -visibility Private
        new-variable -name edge_exe_path -value $null -Description 'EDGEのパス' -Option Private -visibility Private
        new-variable -name process       -value $null -Description 'EDGEのプロセス' -Option Private -visibility Private
        new-variable -name process_id    -value $null -Description 'EDGEのプロセスID' -Option Private -visibility Private
    }

    process
    {
        try
        {
            #Do this if no exceptions happen
            write-host 'try'
            $edge_exe_path = Get-ItemPropertyValue -Path $EDGE_EXE_REG_KEY -Name '(default)'
            write-host '$edge_exe_path:'$edge_exe_path

            # ディレクトリを作成
            if (-not (Test-Path -Path 'C:\temp\UserDataDirectory\'))
            {
                New-Item -Path 'C:\temp\UserDataDirectory\' -Force -ItemType Directory | Out-Null
            }

            # Edgeをデバッグモードで開く
            $process = Start-Process -FilePath $edge_exe_path -ArgumentList '--remote-debugging-port=9222 --disable-popup-blocking --no-first-run --disable-fre --user-data-dir=C:\temp\UserDataDirectory\' -PassThru
                # --remote-debugging-port=9222: デバッグ用 WebSocket をポート 9222 で有効化。
                # --disable-popup-blocking: パップアップを無効化。
                # --no-first-run: 最初の起動を無効化。
                # --disable-fre: フリを無効化。
                # --user-data-dir: ユーザーデータを指定。
            $process_id = $process.Id
            write-host '$process_id:'$process_id
            Start-Sleep -Seconds 2
            # ウィンドウをアクティブにする
            [Microsoft.VisualBasic.Interaction]::AppActivate($process.Id)

            # デバッグ対象のWebSocket URLを取得
            $tabs = Invoke-RestMethod -Uri 'http://localhost:9222/json' -Erroraction Stop
            if (-not $tabs) { throw 'タブ情報を取得できません。' }
            $tab = $tabs[0]  # 最初のタブを選択
            $debugger_url = $tab.webSocketDebuggerUrl
            write-host '$debugger_url:'$debugger_url

            # WebSocket接続の準備
            $web_socket = [System.Net.WebSockets.ClientWebSocket]::new()
            $uri = [System.Uri]::new($debugger_url)
            $web_socket.ConnectAsync($uri, [System.Threading.CancellationToken]::None).Wait()


            # CDPコマンド送信
            $id = $id + 1
            $method = 'Target.setDiscoverTargets'
            $command = @{
                id = $id
                method = $method
                params = @{
                    discover = $true
                }
            } | ConvertTo-Json -Compress
            # WebSocketメッセージ送信
            SendWebSocketMessage -web_socket $web_socket -message $command
            # WebSocketメッセージ受信
            ReceiveWebSocketMessage -web_socket $web_socket



            # CDPコマンド送信
            $id = $id + 1
            $method = 'Page.enable'
            $command = @{
                id = $id
                method = $method
            } | ConvertTo-Json -Compress
            # WebSocketメッセージ送信
            SendWebSocketMessage -web_socket $web_socket -message $command
            # WebSocketメッセージ受信
            ReceiveWebSocketMessage -web_socket $web_socket

            # CDPコマンド送信
            $id = $id + 1
            $url = 'https://www.google.com/'
            $method = 'Page.navigate'
            $command = @{
                id = $id
                method = $method
                params = @{
                    url = $url
                }                
            } | ConvertTo-Json -Compress
            # WebSocketメッセージ送信
            SendWebSocketMessage -web_socket $web_socket -message $command
            # WebSocketメッセージ受信
            ReceiveWebSocketMessage -web_socket $web_socket

            # WebSocket切断
            $web_socket.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, 'Closing', [System.Threading.CancellationToken]::None).Wait()

            # Edgeを終了
            Stop-Process -Id $process_id

            $result = 0
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
        remove-variable -name 'EDGE_EXE_REG_KEY' -ErrorAction SilentlyContinue
        remove-variable -name 'result'           -ErrorAction SilentlyContinue
        remove-variable -name 'edge_exe_path'    -ErrorAction SilentlyContinue
        remove-variable -name 'process'          -ErrorAction SilentlyContinue    
        remove-variable -name 'process_id'       -ErrorAction SilentlyContinue
    }
}

# WebSocketメッセージ送信
function SendWebSocketMessage
{
    param
    (
        [Parameter(Mandatory=$true)]
        [System.Net.WebSockets.ClientWebSocket]$web_socket,

        [Parameter(Mandatory=$true)]
        [string]$message        
    )
    try
    {
        write-host 'SendWebSocketMessage_try'
        write-host '>>>>>>送信内容:'$message
        # WebSocketメッセージ送信
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($message)
        $buffer = [System.ArraySegment[byte]]::new($bytes)
        $web_socket.SendAsync($buffer, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [System.Threading.CancellationToken]::None).Wait()
    }
    catch
    {
        write-host 'SendWebSocketMessage_catch'
        write-host $_.Exception.Message
    }
}

# WebSocketメッセージ受信
function ReceiveWebSocketMessage
{
    param
    (
        [Parameter(Mandatory=$true)]
        [System.Net.WebSockets.ClientWebSocket]$web_socket
    )
    $buffer = New-Object byte[] 4096
    $segment = [System.ArraySegment[byte]]::new($buffer)
    $receive_task = $web_socket.ReceiveAsync($segment, [System.Threading.CancellationToken]::None)
    $receive_task.Wait()
    $response_json = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $receive_task.Result.Count)
    write-host '<<<<<<受信内容:'$response_json
    return $response_json
}

$result = Main
if ($result -eq 0)
{
    exit 0
}
else
{
    exit 1
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
