#Requires -Version 5.1
<#
.SYNOPSIS
    校园网自动登录 —— 菜单 / 托盘 / 开机自启 / 后台守护 / 多账号 + 内存清理
    支持中国电信 Dr.COM、广东电信天翼校园网,以及可配置的通用表单门户

    版本说明:v1.2.4(精简版)—— 与 v1.2.4 完整版同版本号,
    但本版只含「校园网 + 内存清理」,不含桌面宠物 / 图标互动 / 天气。

.DESCRIPTION
    默认按 Dr.COM 门户走(逆向自门户页面 logic_new.js):

      POST http://1.1.1.4/ac_portal/login.php
      opr=pwdLogin & userName=<账号> & pwd=<RC4 hex> & auth_tag=<毫秒时间戳> & rememberPwd=0
      密码用 RC4 加密,密钥 = auth_tag 的十进制字符串,输出小写 hex

    因为 auth_tag 依赖系统时间,脚本会用门户服务器返回的 Date 头自动校正本机时间。

    换个学校就换协议:门户参数都在 portal.json 里,`-Mode portal` 查看/切换。
    已经内置:
      drcom     Dr.COM(RC4 + auth_tag)
      esurfing  广东电信天翼校园网(enet.10000.gd.cn 客户端协议,JSON + 大写 MD5 签名)
      form      通用表单(明文 / md5 / sha1 / base64)
      srun      深澜 Srun(骨架)

.PARAMETER Mode
    menu | setup | switch | delete | list | login | status | watch | tray | install | uninstall | synctime
    wifi | doctor | check | background | log | repair | notify | nicreset | memclean | portal | collect | wizard

.NOTES
#>
[CmdletBinding()]
param(
    [ValidateSet('menu', 'setup', 'switch', 'delete', 'list', 'login', 'status', 'watch', 'tray', 'install', 'uninstall', 'synctime', 'wifi', 'doctor', 'check', 'background', 'log', 'repair', 'notify', 'nicreset', 'memclean', 'portal', 'collect', 'wizard')]
    [string]$Mode = 'menu',

    [string]$User,
    [string]$Password,
    [string]$Label,
    [string]$PortalHost = '1.1.1.4',
    [string]$PortalPage = '/ac_portal/20250303113805/pc.html?template=20250303113805&tabs=pwd&dual_stack=0&vlanid=1033',
    [int]$IntervalSec = 30,

    [switch]$Background,
    [switch]$MachineScope,
    [switch]$Force,
    [switch]$NoVbs,
    [switch]$NoTray,
    [switch]$Elevated,
    [string]$WifiSsid = '',
    [switch]$NoWifi,
    [ValidateSet('light', 'system', 'all')]
    [string]$Scope = 'all',
    [switch]$Deep,
    [string]$BackgroundImage,
    [switch]$ClearBackground,
    [double]$BackgroundOpacity = 0.35,

    # 诊断/自动修复
    [switch]$NoRepair,
    [ValidateSet('safe', 'full')]
    [string]$RepairLevel = 'safe',

    # 由托盘在休眠唤醒时唤起的一次性登录
    [switch]$Resume,


    # 手动指定网卡(一般不用给,默认自动找无线网卡)
    [string]$NicName,

    # 门户配置(通用认证)
    [string]$PortalId,
    [switch]$DetectPortal,
    [switch]$AutoPortal,


    # 换一组探测地址(默认那几个外网地址;调试/内网测试时用)
    [string[]]$ProbeUrl,

    # 兼容旧参数
    [switch]$Status,
    [switch]$Watch
)

$ErrorActionPreference = 'Stop'

if ($Status) { $Mode = 'status' }
if ($Watch) { $Mode = 'watch' }

# ---------------- 路径与常量 ----------------
$WorkDir    = $PSScriptRoot
$SelfScript = Join-Path $WorkDir 'Connect-CampusNet.ps1'
$CookieJar  = Join-Path $WorkDir 'cookies.txt'
$AccountsFn = Join-Path $WorkDir 'accounts.json'
$SettingsFn = Join-Path $WorkDir 'settings.json'
$LegacyFn   = Join-Path $WorkDir 'cred.json'
$StateFn    = Join-Path $WorkDir 'state.json'
$LogFile    = Join-Path $WorkDir 'campus-net.log'
$VbsFile    = Join-Path $WorkDir 'run-hidden.vbs'
$TaskName   = 'CampusNet-AutoLogin'
$RepairTaskName = 'CampusNet-Repair'
$MemCleanTaskName = 'CampusNet-MemClean'
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
$script:AppVersion = '1.2.4'
$script:AppEdition = '精简版'   # 版本号 1.2.4,但本版不含桌面宠物,只有校园网 + 内存清理
$Curl = (Get-Command curl.exe -ErrorAction SilentlyContinue).Source
$script:NicList = $null
$script:NicListAt = [datetime]::MinValue
$script:TrayAlive = $false
$script:TrayAliveAt = [datetime]::MinValue

# ---------------- 设置(可保存的开关) ----------------
function Get-Settings {
    $def = [pscustomobject]@{
        autoWifi        = $true
        wifiSsid        = 'CIT-Chinanet'
        lastPortalCheck = ''
        lastRepair      = ''
        lastNicReset    = ''
        lastResume      = ''
        memCleanAuto    = $true
        memCleanMinFreePct = 12
        memCleanCooldownMin = 10
        memCleanSysInAuto = $false
        lastMemClean    = ''
    }
    if (-not (Test-Path $SettingsFn)) { return $def }
    try {
        $s = Get-Content $SettingsFn -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -eq $s.autoWifi) { $s | Add-Member -NotePropertyName autoWifi -NotePropertyValue $def.autoWifi -Force }
        if (-not $s.wifiSsid) { $s | Add-Member -NotePropertyName wifiSsid -NotePropertyValue $def.wifiSsid -Force }
        if (-not $s.PSObject.Properties['lastPortalCheck']) { $s | Add-Member -NotePropertyName lastPortalCheck -NotePropertyValue '' -Force }
        if (-not $s.PSObject.Properties['lastRepair']) { $s | Add-Member -NotePropertyName lastRepair -NotePropertyValue '' -Force }
        if (-not $s.PSObject.Properties['lastNicReset']) { $s | Add-Member -NotePropertyName lastNicReset -NotePropertyValue '' -Force }
        if (-not $s.PSObject.Properties['lastResume']) { $s | Add-Member -NotePropertyName lastResume -NotePropertyValue '' -Force }
        if (-not $s.PSObject.Properties['memCleanAuto']) { $s | Add-Member -NotePropertyName memCleanAuto -NotePropertyValue $true -Force }
        if (-not $s.PSObject.Properties['memCleanMinFreePct']) { $s | Add-Member -NotePropertyName memCleanMinFreePct -NotePropertyValue 12 -Force }
        if (-not $s.PSObject.Properties['memCleanCooldownMin']) { $s | Add-Member -NotePropertyName memCleanCooldownMin -NotePropertyValue 10 -Force }
        if (-not $s.PSObject.Properties['memCleanSysInAuto']) { $s | Add-Member -NotePropertyName memCleanSysInAuto -NotePropertyValue $false -Force }
        if (-not $s.PSObject.Properties['lastMemClean']) { $s | Add-Member -NotePropertyName lastMemClean -NotePropertyValue '' -Force }
        return $s
    } catch { return $def }
}

function Save-Settings {
    param($Settings)
    [pscustomobject]@{
        autoWifi        = [bool]$Settings.autoWifi
        wifiSsid        = [string]$Settings.wifiSsid
        lastPortalCheck = [string]$Settings.lastPortalCheck
        lastRepair      = [string]$Settings.lastRepair
        lastNicReset    = [string]$Settings.lastNicReset
        lastResume      = [string]$Settings.lastResume
        memCleanAuto    = [bool]$(if ($null -ne $Settings.memCleanAuto) { $Settings.memCleanAuto } else { $true })
        memCleanMinFreePct = [int]$(if ($Settings.memCleanMinFreePct) { $Settings.memCleanMinFreePct } else { 12 })
        memCleanCooldownMin = [int]$(if ($Settings.memCleanCooldownMin) { $Settings.memCleanCooldownMin } else { 10 })
        memCleanSysInAuto = [bool]$(if ($null -ne $Settings.memCleanSysInAuto) { $Settings.memCleanSysInAuto } else { $false })
        lastMemClean    = [string]$Settings.lastMemClean
    } | ConvertTo-Json | Set-Content -Path $SettingsFn -Encoding UTF8
}

# 命令行 -WifiSsid 属于一次性指定:给了就写进设置,之后统一以设置为准
$script:Settings = Get-Settings
if ($WifiSsid -and $WifiSsid -ne $script:Settings.wifiSsid) {
    $script:Settings.wifiSsid = $WifiSsid
    Save-Settings -Settings $script:Settings
}
$WifiSsid = $script:Settings.wifiSsid
$script:AutoWifi = [bool]$script:Settings.autoWifi

# ---------------- 日志 ----------------
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message

    if (-not $Background) {
        switch ($Level) {
            'OK'    { Write-Host $line -ForegroundColor Green }
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'ERROR' { Write-Host $line -ForegroundColor Red }
            default { Write-Host $line }
        }
    }
    try {
        if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt 512KB) {
            Move-Item $LogFile ($LogFile + '.old') -Force
        }
        Add-Content -Path $LogFile -Value $line -Encoding UTF8
    } catch { }
}

function Write-Stdout { param([string]$Message) Write-Host $Message }

# ================= 内存清理(Mem Reduct 精简版)=================
# 只做三件事:① 自家进程 EmptyWorkingSet ② 清 standby list ③ 清系统文件缓存
# ②③ 需要管理员;免 UAC 的方式是由最高权限计划任务代为执行(见 Register-MemCleanTask)

$script:MemCs = @'
using System;
using System.Runtime.InteropServices;

namespace CampusNetMem {
    public static class Mem {
        [StructLayout(LayoutKind.Sequential)]
        public struct MEMORYSTATUSEX {
            public uint dwLength;
            public uint dwMemoryLoad;
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool GlobalMemoryStatusEx(ref MEMORYSTATUSEX lpBuffer);

        [DllImport("psapi.dll", SetLastError = true)]
        static extern bool EmptyWorkingSet(IntPtr hProcess);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr OpenProcess(int access, bool inherit, int pid);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool CloseHandle(IntPtr h);

        [DllImport("ntdll.dll")]
        static extern int NtSetSystemInformation(int infoClass, IntPtr info, int length);

        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool OpenProcessToken(IntPtr h, uint access, out IntPtr token);

        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool LookupPrivilegeValue(string system, string name, out long luid);

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        struct TOKEN_PRIVILEGES {
            public uint Count;
            public long Luid;
            public uint Attr;
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool AdjustTokenPrivileges(IntPtr token, bool disableAll, ref TOKEN_PRIVILEGES newState, int bufferLength, IntPtr prevState, IntPtr returnLength);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetCurrentProcess();

        // 返回:总物理内存 / 可用物理内存 / 内存负载百分比
        public static ulong[] Snapshot() {
            MEMORYSTATUSEX m = new MEMORYSTATUSEX();
            m.dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX));
            GlobalMemoryStatusEx(ref m);
            return new ulong[] { m.ullTotalPhys, m.ullAvailPhys, m.dwMemoryLoad };
        }

        // 把某个进程的工作集换出(用户态即可,不需要管理员)
        public static bool TrimProcess(int pid) {
            const int PROCESS_SET_QUOTA = 0x0100;
            const int PROCESS_QUERY_INFORMATION = 0x0400;
            IntPtr h = OpenProcess(PROCESS_SET_QUOTA | PROCESS_QUERY_INFORMATION, false, pid);
            if (h == IntPtr.Zero) return false;
            try { return EmptyWorkingSet(h); }
            finally { CloseHandle(h); }
        }

        static bool EnablePrivilege(string name) {
            const uint TOKEN_ADJUST_PRIVILEGES = 0x0020;
            const uint TOKEN_QUERY = 0x0008;
            const uint SE_PRIVILEGE_ENABLED = 0x0002;
            IntPtr token;
            if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, out token)) return false;
            try {
                long luid;
                if (!LookupPrivilegeValue(null, name, out luid)) return false;
                TOKEN_PRIVILEGES tp = new TOKEN_PRIVILEGES();
                tp.Count = 1;
                tp.Luid = luid;
                tp.Attr = SE_PRIVILEGE_ENABLED;
                return AdjustTokenPrivileges(token, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            } finally { CloseHandle(token); }
        }

        // SystemMemoryListInformation = 0x50,命令 4 = PurgeStandbyList
        public static int PurgeStandby() {
            EnablePrivilege("SeProfileSingleProcessPrivilege");
            EnablePrivilege("SeIncreaseQuotaPrivilege");
            IntPtr p = Marshal.AllocHGlobal(4);
            try {
                Marshal.WriteInt32(p, 4);
                return NtSetSystemInformation(0x50, p, 4);
            } finally { Marshal.FreeHGlobal(p); }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct SYSTEM_FILECACHE_INFORMATION {
            public IntPtr CurrentSize;
            public IntPtr PeakSize;
            public IntPtr PageFaultCount;
            public IntPtr MinimumWorkingSet;
            public IntPtr MaximumWorkingSet;
            public IntPtr CurrentSizeIncludingTransitionInPages;
            public IntPtr PeakSizeIncludingTransitionInPages;
            public uint TransitionRePurposeCount;
            public uint Flags;
        }

        // SystemFileCacheInformation = 0x15;Minimum/MaximumWorkingSet = -1 让系统回收文件缓存
        public static int ClearFileCache() {
            EnablePrivilege("SeIncreaseQuotaPrivilege");
            SYSTEM_FILECACHE_INFORMATION info = new SYSTEM_FILECACHE_INFORMATION();
            info.MinimumWorkingSet = (IntPtr)(-1);
            info.MaximumWorkingSet = (IntPtr)(-1);
            int size = Marshal.SizeOf(typeof(SYSTEM_FILECACHE_INFORMATION));
            IntPtr p = Marshal.AllocHGlobal(size);
            try {
                Marshal.StructureToPtr(info, p, false);
                return NtSetSystemInformation(0x15, p, size);
            } finally { Marshal.FreeHGlobal(p); }
        }
    }
}
'@

$script:MemApiReady = $false

function Initialize-MemApi {
    if ($script:MemApiReady) { return $true }
    try {
        if (-not ('CampusNetMem.Mem' -as [type])) {
            Add-Type -TypeDefinition $script:MemCs -Language CSharp
        }
        $script:MemApiReady = $true
        return $true
    } catch {
        Write-Log ('内存接口加载失败:' + $_.Exception.Message) 'WARN'
        return $false
    }
}

function Get-MemorySnapshot {
    # 统一的内存快照:总 MB / 可用 MB / 可用百分比 / 负载百分比
    $r = @{ TotalMB = 0; FreeMB = 0; FreePct = 0; LoadPct = 0 }
    try {
        if (-not (Initialize-MemApi)) { return $r }
        $s = [CampusNetMem.Mem]::Snapshot()
        $r.TotalMB = [int]($s[0] / 1MB)
        $r.FreeMB = [int]($s[1] / 1MB)
        $r.LoadPct = [int]$s[2]
        if ($r.TotalMB -gt 0) { $r.FreePct = [int][Math]::Round(100.0 * $r.FreeMB / $r.TotalMB) }
    } catch { }
    return $r
}

function Get-CampusNetSelfPids {
    $list = @()
    try {
        $procs = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue
        foreach ($p in $procs) {
            if ($p.ProcessId -eq $PID) { continue }
            $cmd = [string]$p.CommandLine
            if ($cmd -notmatch 'Connect-CampusNet\.ps1') { continue }
        }
    } catch { }
    return $list
}

function Get-SelfWorkingSetMB {
    param([int[]]$Pids)
    $sum = 0
    foreach ($p in $Pids) {
        try { $sum += [int]((Get-Process -Id $p -ErrorAction SilentlyContinue).WorkingSet64 / 1MB) } catch { }
    }
    return $sum
}

function Invoke-MemoryCleanLight {
    # 轻清:自家进程工作集换出 + GC。零权限,随时可做
    $before = Get-MemorySnapshot
    $pids = @(Get-CampusNetSelfPids)
    $wsBefore = Get-SelfWorkingSetMB -Pids $pids
    $done = 0
    foreach ($p in $pids) {
        try { if ([CampusNetMem.Mem]::TrimProcess($p)) { $done++ } } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $wsAfter = Get-SelfWorkingSetMB -Pids $pids
    $after = Get-MemorySnapshot
    $msg = '内存轻清:' + $done + '/' + $pids.Count + ' 个自家进程,工作集 ' + $wsBefore + '→' + $wsAfter + ' MB,可用 ' + $before.FreeMB + '→' + $after.FreeMB + ' MB'
    Write-Log $msg 'OK'
    return @{ Ok = ($done -gt 0); Message = $msg; FreeBefore = $before.FreeMB; FreeAfter = $after.FreeMB; WsBefore = $wsBefore; WsAfter = $wsAfter }
}

function Invoke-MemoryCleanSystem {
    # 系统清:清 standby list + 文件缓存(需管理员;非管理员会拿到 NTSTATUS 失败码)
    $before = Get-MemorySnapshot
    if (-not (Initialize-MemApi)) { return @{ Ok = $false; Message = '内存接口不可用' } }
    $st1 = 0
    $st2 = 0
    try { $st1 = [CampusNetMem.Mem]::PurgeStandby() } catch { $st1 = -99 }
    try { $st2 = [CampusNetMem.Mem]::ClearFileCache() } catch { $st2 = -99 }
    Start-Sleep -Milliseconds 300
    $after = Get-MemorySnapshot
    $ok = ($st1 -eq 0)
    $msg = '内存系统清理:standby ' + $(if ($st1 -eq 0) { 'OK' } else { '失败(0x{0:X8})' -f $st1 }) + ',文件缓存 ' + $(if ($st2 -eq 0) { 'OK' } else { '失败(0x{0:X8})' -f $st2 }) + ',可用 ' + $before.FreeMB + '→' + $after.FreeMB + ' MB'
    if (-not $ok) { $msg += ' —— 需要管理员权限' }
    Write-Log $msg $(if ($ok) { 'OK' } else { 'WARN' })
    return @{ Ok = $ok; Message = $msg; FreeBefore = $before.FreeMB; FreeAfter = $after.FreeMB }
}

function Invoke-MemoryClean {
    param([ValidateSet('light', 'system', 'all')][string]$Scope = 'light')
    $parts = @()
    if ($Scope -in @('light', 'all')) { $parts += (Invoke-MemoryCleanLight) }
    if ($Scope -in @('system', 'all')) { $parts += (Invoke-MemoryCleanSystem) }
    $free = (Get-MemorySnapshot).FreeMB
    try {
        $s = Get-Settings
        $s.lastMemClean = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Save-Settings -Settings $s
    } catch { }
    return @{ Parts = $parts; FreeMB = $free }
}

function Get-MemCleanTask {
    try { return (Get-ScheduledTask -TaskName $MemCleanTaskName -ErrorAction SilentlyContinue) } catch { return $null }
}

function Register-MemCleanTask {
    # 提权清理任务:平时不跑,只有托盘点「系统清理」时才被唤起(不弹 UAC)
    if (-not (Test-Elevated)) {
        Write-Log '当前不是管理员:不注册内存清理提权任务,托盘里的「系统清理」将不可用。要用请用管理员运行安装(加 -Elevated)' 'WARN'
        return $false
    }
    try {
        $arg = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -Mode memclean -Scope all -Background' -f $SelfScript
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 3)
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest
        if (Get-ScheduledTask -TaskName $MemCleanTaskName -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $MemCleanTaskName -Confirm:$false
        }
        Register-ScheduledTask -TaskName $MemCleanTaskName -Action $action -Settings $settings `
            -Principal $principal -Description '校园网自动登录:提权清理内存(供托盘调用,不弹 UAC)' | Out-Null
        Write-Log ('已注册提权内存清理任务 ' + $MemCleanTaskName) 'OK'
        return $true
    } catch {
        Write-Log ('注册提权内存清理任务失败:' + $_.Exception.Message) 'WARN'
        return $false
    }
}

function Start-MemCleanTask {
    # 用提权任务做系统清理(不弹 UAC);任务不存在时返回 false
    $task = Get-MemCleanTask
    if (-not $task) { return $false }
    try {
        Start-ScheduledTask -TaskName $MemCleanTaskName
        return $true
    } catch {
        Write-Log ('唤起内存清理任务失败:' + $_.Exception.Message) 'WARN'
        return $false
    }
}

function Test-MemCleanDue {
    # 自动清理判断:开关 + 阈值(可用百分比) + 冷却
    try {
        $s = Get-Settings
        if (-not $s.memCleanAuto) { return $false }
        $last = [string]$s.lastMemClean
        if ($last) {
            try {
                $mins = ((Get-Date) - [datetime]::Parse($last)).TotalMinutes
                if ($mins -lt [int]$s.memCleanCooldownMin) { return $false }
            } catch { }
        }
        $m = Get-MemorySnapshot
        if ($m.TotalMB -le 0) { return $false }
        return ($m.FreePct -lt [int]$s.memCleanMinFreePct)
    } catch { return $false }
}

# ---------------- RC4(照抄门户 JS,已交叉验证) ----------------
function Get-Rc4Hex {
    param([string]$Text, [string]$Key)

    $keyBytes = New-Object int[] 256
    for ($i = 0; $i -lt 256; $i++) { $keyBytes[$i] = [int][char]$Key[$i % $Key.Length] }

    $sbox = 0..255
    $j = 0
    for ($i = 0; $i -lt 256; $i++) {
        $j = ($j + $sbox[$i] + $keyBytes[$i]) % 256
        $t = $sbox[$i]; $sbox[$i] = $sbox[$j]; $sbox[$j] = $t
    }

    $a = 0; $b = 0
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $a = ($a + 1) % 256
        $b = ($b + $sbox[$a]) % 256
        $t = $sbox[$a]; $sbox[$a] = $sbox[$b]; $sbox[$b] = $t
        $c = ($sbox[$a] + $sbox[$b]) % 256
        $x = ([int][char]$Text[$i]) -bxor $sbox[$c]
        $hex = $x.ToString('x')
        if ($hex.Length -eq 1) { $hex = '0' + $hex }
        [void]$sb.Append($hex)
    }
    $sb.ToString()
}

# ---------------- 网速监控(常驻,不需要开关) ----------------
# 直接读 .NET 的网卡字节计数:不起进程、不查 WMI、不碰注册表,
# 2 秒采样一次的开销可以忽略,所以常开也不心疼。
$script:SpeedLast   = $null
$script:SpeedResult = $null
$script:SpeedPeakD  = [double]0
$script:SpeedPeakU  = [double]0
$script:SpeedTotalD = [double]0
$script:SpeedTotalU = [double]0

function Get-NicByteCounters {
    # 累加所有「已启用 + 非回环 + 非隧道」网卡的字节数。
    # GetAllNetworkInterfaces() 要十几毫秒,而网卡列表几乎不变 —— 缓存 30 秒,
    # 每次只重读统计值,这样托盘每 2 秒的网速刷新不会被别的活顶掉。
    $rx = [long]0
    $tx = [long]0
    try {
        $now = Get-Date
        if (-not $script:NicList -or ($now - $script:NicListAt).TotalSeconds -ge 30) {
            $list = New-Object System.Collections.ArrayList
            foreach ($ni in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
                try {
                    if ($ni.OperationalStatus -ne [System.Net.NetworkInformation.OperationalStatus]::Up) { continue }
                    $t = $ni.NetworkInterfaceType
                    if ($t -eq [System.Net.NetworkInformation.NetworkInterfaceType]::Loopback) { continue }
                    if ($t -eq [System.Net.NetworkInformation.NetworkInterfaceType]::Tunnel) { continue }
                    [void]$list.Add($ni)
                } catch { }
            }
            $script:NicList = $list
            $script:NicListAt = $now
        }
        foreach ($ni in $script:NicList) {
            try {
                $st = $ni.GetIPv4Statistics()
                if ($st) { $rx += $st.BytesReceived; $tx += $st.BytesSent }
            } catch { }
        }
    } catch { }
    return [pscustomobject]@{ Rx = $rx; Tx = $tx }
}

function Format-Speed {
    param([double]$BytesPerSec, [switch]$Short)
    if ($BytesPerSec -lt 0) { $BytesPerSec = 0 }
    if ($Short) {
        if ($BytesPerSec -lt 1024) { return ('{0:N0}B' -f $BytesPerSec) }
        if ($BytesPerSec -lt 1048576) { return ('{0:N0}K' -f ($BytesPerSec / 1024)) }
        return ('{0:N1}M' -f ($BytesPerSec / 1048576))
    }
    if ($BytesPerSec -lt 1024) { return ('{0:N0} B/s' -f $BytesPerSec) }
    if ($BytesPerSec -lt 1048576) { return ('{0:N0} KB/s' -f ($BytesPerSec / 1024)) }
    if ($BytesPerSec -lt 1073741824) { return ('{0:N2} MB/s' -f ($BytesPerSec / 1048576)) }
    return ('{0:N2} GB/s' -f ($BytesPerSec / 1073741824))
}

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -lt 1024) { return ('{0:N0} B' -f $Bytes) }
    if ($Bytes -lt 1048576) { return ('{0:N0} KB' -f ($Bytes / 1024)) }
    if ($Bytes -lt 1073741824) { return ('{0:N1} MB' -f ($Bytes / 1048576)) }
    return ('{0:N2} GB' -f ($Bytes / 1073741824))
}

function Get-NetSpeed {
    param([switch]$Reset)

    if ($Reset) {
        $script:SpeedLast = $null; $script:SpeedResult = $null
        $script:SpeedPeakD = 0; $script:SpeedPeakU = 0
        $script:SpeedTotalD = 0; $script:SpeedTotalU = 0
        return $null
    }

    $now = Get-Date
    $c = Get-NicByteCounters

    # 第一次只打基线,没法算速度(否则会拿"开机以来的平均速度"糊弄人)
    if (-not $script:SpeedLast) {
        $script:SpeedLast = [pscustomobject]@{ Time = $now; Rx = $c.Rx; Tx = $c.Tx }
        $script:SpeedResult = [pscustomobject]@{
            Ready = $false; Down = [double]0; Up = [double]0
            DownText = '—'; UpText = '—'; DownShort = '—'; UpShort = '—'
            PeakDownText = '—'; PeakUpText = '—'; TotalText = '—'
        }
        return $script:SpeedResult
    }

    $el = ($now - $script:SpeedLast.Time).TotalSeconds
    if ($el -lt 0.5 -and $script:SpeedResult) { return $script:SpeedResult }
    if ($el -le 0) { $el = 0.5 }

    $dRx = $c.Rx - $script:SpeedLast.Rx
    $dTx = $c.Tx - $script:SpeedLast.Tx
    # 网卡重连/计数器归零会出现负增量,当 0 处理,别算出天文数字
    if ($dRx -lt 0 -or $dTx -lt 0) { $dRx = 0; $dTx = 0 }

    $down = [double]$dRx / $el
    $up = [double]$dTx / $el

    $script:SpeedLast = [pscustomobject]@{ Time = $now; Rx = $c.Rx; Tx = $c.Tx }
    $script:SpeedTotalD += $dRx
    $script:SpeedTotalU += $dTx
    if ($down -gt $script:SpeedPeakD) { $script:SpeedPeakD = $down }
    if ($up -gt $script:SpeedPeakU) { $script:SpeedPeakU = $up }

    $script:SpeedResult = [pscustomobject]@{
        Ready        = $true
        Down         = $down
        Up           = $up
        DownText     = (Format-Speed $down)
        UpText       = (Format-Speed $up)
        DownShort    = (Format-Speed $down -Short)
        UpShort      = (Format-Speed $up -Short)
        PeakDownText = (Format-Speed $script:SpeedPeakD)
        PeakUpText   = (Format-Speed $script:SpeedPeakU)
        TotalText    = ('↓' + (Format-Bytes $script:SpeedTotalD) + '  ↑' + (Format-Bytes $script:SpeedTotalU))
    }
    return $script:SpeedResult
}

# ---------------- WiFi ----------------
# 两个坑要一起处理:
#  1) Windows 11 26H2 起,netsh 的 WLAN 查询命令要求“位置权限”,没开就只回一句
#     “网络 shell 命令需要位置权限才能访问 WLAN 信息” + 错误 5。
#  2) 26H2 的 netsh 输出是 UTF-8,而脚本常跑在没有控制台/输出被重定向的宿主里,
#     那时 [Console]::OutputEncoding 是 GBK,直接 & 调用会把中文解成乱码写进日志。
# 所以统一用 Invoke-NativeLines 读原始字节再智能解码,并保留退出码。
$script:WlanLocationBlocked = $false
$script:WlanRepairAt         = [datetime]::MinValue
$script:NativeLastExitCode  = 0

function Invoke-NativeLines {
    param([string]$FilePath, [string]$Arguments = '')

    $script:NativeLastExitCode = 0
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $FilePath
    $psi.Arguments              = $Arguments
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    $ms  = New-Object System.IO.MemoryStream
    $buf = New-Object byte[] 8192
    $s   = $proc.StandardOutput.BaseStream
    while ($true) {
        $n = $s.Read($buf, 0, $buf.Length)
        if ($n -le 0) { break }
        $ms.Write($buf, 0, $n)
    }
    $errText = ''
    try { $errText = $proc.StandardError.ReadToEnd() } catch { }
    [void]$proc.WaitForExit(20000)
    try { $script:NativeLastExitCode = $proc.ExitCode } catch { }
    $proc.Dispose()

    $bytes = $ms.ToArray()
    $ms.Dispose()
    $body = ''
    if ($bytes.Length -gt 0) {
        try {
            $strict = New-Object System.Text.UTF8Encoding($false, $true)
            $body = $strict.GetString($bytes)
        } catch {
            $body = [System.Text.Encoding]::GetEncoding(936).GetString($bytes)
        }
    }
    if ($errText) { $body = $body + "`r`n" + $errText }
    return @($body -split "`r?`n")
}

function Get-NetshPath {
    return (Join-Path $env:SystemRoot 'System32\netsh.exe')
}

function Test-WlanLocationBlocked {
    param([string[]]$Lines)
    return [bool](($Lines -join "`n") -match '位置权限|privacy-location|error 5|错误 5|access is denied')
}

function Get-WifiInfo {
    $ssid = ''
    $state = ''
    $ifName = ''
    $out = @()
    try { $out = Invoke-NativeLines (Get-NetshPath) 'wlan show interfaces' } catch { $out = @() }

    $blocked = Test-WlanLocationBlocked $out
    if ($blocked) { $script:WlanLocationBlocked = $true }
    elseif (($out -join "`n") -match '(?m)^\s*SSID\b|(?m)^\s*(State|状态)\s*:') { $script:WlanLocationBlocked = $false }

    foreach ($line in $out) {
        if ($line -match '^\s*(Name|名称)\s*:\s*(.+)$' -and -not $ifName) { $ifName = $Matches[2].Trim() }
        elseif ($line -match '^\s*SSID\s*:\s*(.+)$') { $ssid = $Matches[1].Trim() }
        elseif ($line -match '^\s*(State|状态)\s*:\s*(.+)$') { $state = $Matches[2].Trim() }
    }

    # netsh 被位置权限挡住时,退回 NLA(Get-NetAdapter / Get-NetConnectionProfile),
    # 这条路不经过位置权限,能可靠判断“连没连、连的是哪个 SSID”。
    if (-not $state -or -not $ssid) {
        try {
            $ad = @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.PhysicalMediaType -match '802\.11' })
            if ($ad.Count -gt 0) {
                if (-not $ifName) { $ifName = $ad[0].Name }
                if ($ad[0].Status -eq 'Up') {
                    $p = @(Get-NetConnectionProfile -InterfaceIndex $ad[0].ifIndex -ErrorAction SilentlyContinue)
                    if ($p.Count -gt 0) {
                        if (-not $ssid) { $ssid = $p[0].Name }
                        $state = 'connected'
                    } else { $state = 'disconnected' }
                } else { $state = 'disconnected' }
            }
        } catch { }
    }

    return [pscustomobject]@{ State = $state; Ssid = $ssid; Interface = $ifName; LocationBlocked = $script:WlanLocationBlocked }
}

function Test-WifiConnected {
    param([string]$Ssid)

    $info = Get-WifiInfo
    if ($info.State -notmatch '^(connected|已连接)$') { return $false }
    if ($Ssid -and $info.Ssid -ne $Ssid) { return $false }
    return $true
}

# ---------------- WiFi 列表(自动刷新,不需要开关) ----------------
# netsh 读的是系统缓存。每次连接前先跑一遍 show networks 让无线服务重新枚举,
# 再读 profiles / interfaces,免得「明明有信号却报没有配置」。
$script:WifiList     = $null
$script:WifiListTime = [datetime]::MinValue

function Get-WifiProfiles {
    $list = New-Object System.Collections.ArrayList
    try {
        foreach ($line in (Invoke-NativeLines (Get-NetshPath) 'wlan show profiles')) {
            if ($line -match ':\s*(\S.*?)\s*$') {
                $n = $Matches[1]
                if ($n -and -not $list.Contains($n)) { [void]$list.Add($n) }
            }
        }
    } catch { }
    return @($list)
}

function Get-WifiVisible {
    # 这次调用本身会让无线服务重新扫一遍;返回 @{ Ssid, Open, Signal }
    $list = New-Object System.Collections.ArrayList
    $cur = $null
    try {
        foreach ($line in (Invoke-NativeLines (Get-NetshPath) 'wlan show networks mode=bssid')) {
            if ($line -match '^SSID\s+\d+\s*:\s*(.*)$') {
                if ($cur) { [void]$list.Add($cur) }
                $cur = [pscustomobject]@{ Ssid = $Matches[1].Trim(); Open = $false; Signal = 0 }
                continue
            }
            if (-not $cur) { continue }
            if ($line -match '^\s*(Authentication|身份验证)\s*:\s*(.+)$') {
                if ($Matches[2].Trim() -match '^(Open|开放式|开放|None)') { $cur.Open = $true }
            } elseif ($line -match '^\s*(Signal|信号)\s*:\s*(\d+)\s*%') {
                $sig = [int]$Matches[2]
                if ($sig -gt $cur.Signal) { $cur.Signal = $sig }
            }
        }
        if ($cur) { [void]$list.Add($cur) }
    } catch { }
    return @($list)
}

function Update-WifiList {
    param([switch]$Force, [int]$MinIntervalSec = 40)

    $now = Get-Date
    if (-not $Force -and $script:WifiList -and (($now - $script:WifiListTime).TotalSeconds -lt $MinIntervalSec)) {
        return $script:WifiList
    }

    $prev     = $script:WifiList
    $visible  = Get-WifiVisible
    $profiles = Get-WifiProfiles
    $info     = Get-WifiInfo

    $tv = $null
    foreach ($v in $visible) { if ($v.Ssid -eq $WifiSsid) { $tv = $v } }

    $res = [pscustomobject]@{
        Time             = $now
        Visible          = $visible
        Profiles         = $profiles
        Current          = $info.Ssid
        Connected        = [bool]($info.State -match '^(connected|已连接)$')
        Target           = $WifiSsid
        TargetVisible    = [bool]$tv
        TargetOpen       = [bool]($tv -and $tv.Open)
        TargetSignal     = $(if ($tv) { [int]$tv.Signal } else { 0 })
        TargetInProfiles = [bool]($profiles -contains $WifiSsid)
        LocationBlocked  = [bool]$info.LocationBlocked
    }

    $script:WifiList = $res
    $script:WifiListTime = $now

    if ($prev -and $prev.TargetVisible -ne $res.TargetVisible) {
        if ($res.TargetVisible) { Write-Log ('WiFi 列表更新: ' + $WifiSsid + ' 重新可见(信号 ' + $res.TargetSignal + '%)') 'OK' }
        else { Write-Log ('WiFi 列表更新: ' + $WifiSsid + ' 从列表里消失了') 'WARN' }
    }
    if ($prev -and -not $prev.TargetInProfiles -and $res.TargetInProfiles) {
        Write-Log ('WiFi 列表更新: 已存在 ' + $WifiSsid + ' 的 Windows 配置') 'OK'
    }
    return $res
}

function ConvertTo-XmlEscaped {
    param([string]$Text)
    if (-not $Text) { return '' }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&apos;')
}

function Add-OpenWifiProfile {
    param([string]$Ssid)
    # 只给开放网络用:配置丢了就自动补一个,省得手动点一次
    $esc = ConvertTo-XmlEscaped $Ssid
    $ascii = $true
    foreach ($ch in $Ssid.ToCharArray()) { if ([int]$ch -lt 32 -or [int]$ch -gt 126) { $ascii = $false; break } }
    if ($ascii) {
        $ssidBlock = '<name>' + $esc + '</name>'
    } else {
        $hex = -join ([System.Text.Encoding]::UTF8.GetBytes($Ssid) | ForEach-Object { $_.ToString('x2') })
        $ssidBlock = '<hex>' + $hex + '</hex><name>' + $esc + '</name>'
    }

    $nl = "`r`n"
    $xml = '<?xml version="1.0"?>' + $nl +
    '<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">' + $nl +
    '  <name>' + $esc + '</name>' + $nl +
    '  <SSIDConfig><SSID>' + $ssidBlock + '</SSID></SSIDConfig>' + $nl +
    '  <connectionType>ESS</connectionType>' + $nl +
    '  <connectionMode>auto</connectionMode>' + $nl +
    '  <MSM><security>' + $nl +
    '    <authEncryption><authentication>open</authentication><encryption>none</encryption><useOneX>false</useOneX></authEncryption>' + $nl +
    '  </security></MSM>' + $nl +
    '</WLANProfile>'

    $tmp = Join-Path $env:TEMP ('wifi-profile-' + [guid]::NewGuid().ToString('N') + '.xml')
    try {
        [System.IO.File]::WriteAllText($tmp, $xml, (New-Object System.Text.UTF8Encoding($false)))
        $r = Invoke-NativeLines (Get-NetshPath) ('wlan add profile filename="{0}" user=current' -f $tmp)
        $msg = (($r | Where-Object { $_ }) -join ' ').Trim()
        if ($script:NativeLastExitCode -eq 0) {
            Write-Log ('已自动创建开放网络配置: ' + $Ssid) 'OK'
            return $true
        }
        Write-Log ('自动创建 WiFi 配置失败(' + $script:NativeLastExitCode + '): ' + $msg) 'WARN'
        return $false
    } catch {
        Write-Log ('自动创建 WiFi 配置出错: ' + $_.Exception.Message) 'WARN'
        return $false
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Repair-WlanLocationPermission {
    # Windows 11 26H2 起,「读取 WLAN 信息」被纳入位置权限:netsh wlan show/connect
    # 统统只回一句“网络 shell 命令需要位置权限”+ 错误 5。升级 26H2 后系统常把
    # 定位服务关掉,于是脚本误判“没连 WiFi”,每 30 秒重连一次刷屏。
    # 有管理员权限就自己把开关打开(注册表 + 重启定位服务);否则只写日志提示手开。
    $now = Get-Date
    if ($script:WlanRepairAt -ne [datetime]::MinValue -and ($now - $script:WlanRepairAt).TotalMinutes -lt 10) { return $false }
    if (-not (Test-Elevated)) {
        Write-Log '需要位置权限:设置 → 隐私和安全性 → 位置,打开「定位服务」(或右键以管理员身份运行脚本让它自动修复)' 'ERROR'
        return $false
    }
    $script:WlanRepairAt = $now
    Write-Log '尝试自动修复「位置权限」(Windows 11 26H2 的 WLAN 限制)…' 'WARN'

    try {
        $cs = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location'
        if (-not (Test-Path $cs)) { New-Item -Path $cs -Force | Out-Null }
        New-ItemProperty -Path $cs -Name 'Value' -Value 'Allow' -PropertyType String -Force | Out-Null
        $np = Join-Path $cs 'NonPackaged'
        if (-not (Test-Path $np)) { New-Item -Path $np -Force | Out-Null }
        New-ItemProperty -Path $np -Name 'Value' -Value 'Allow' -PropertyType String -Force | Out-Null
        $npExe = Join-Path $np 'C:#Windows#System32#netsh.exe'
        if (-not (Test-Path $npExe)) { New-Item -Path $npExe -Force | Out-Null }
        New-ItemProperty -Path $npExe -Name 'Value' -Value 'Allow' -PropertyType String -Force | Out-Null
        $cfg = 'HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration'
        if (Test-Path $cfg) { New-ItemProperty -Path $cfg -Name 'Status' -Value 1 -PropertyType DWord -Force | Out-Null }
    } catch {
        Write-Log ('写位置权限注册表失败: ' + $_.Exception.Message) 'WARN'
    }

    try {
        Restart-Service -Name 'lfsvc' -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        Restart-Service -Name 'WlanSvc' -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
        Write-Log '定位服务已重启,重新读取 WLAN 信息' 'OK'
    } catch {
        Write-Log ('重启定位服务失败: ' + $_.Exception.Message) 'WARN'
        return $false
    }

    # 改完不一定立刻生效(设置页和服务都有缓存),这里验一次,没好就给出明确的手动路径
    $after = @()
    try { $after = Invoke-NativeLines (Get-NetshPath) 'wlan show interfaces' } catch { $after = @() }
    if (Test-WlanLocationBlocked $after) {
        Write-Log '自动打开位置权限没立即生效:打开 设置 → 隐私和安全性 → 位置,把「定位服务」打开(可能要等几秒再刷新)' 'WARN'
        return $false
    }
    $script:WlanLocationBlocked = $false
    Write-Log 'netsh 已能正常读取 WLAN 信息' 'OK'
    return $true
}

function Connect-CampusWifi {
    param([string]$Ssid = 'CIT-Chinanet', [int]$TimeoutSec = 25)

    if (Test-WifiConnected -Ssid $Ssid) {
        Write-Log ('WiFi 已连接:' + $Ssid)
        return $true
    }

    $info = Get-WifiInfo
    if ($info.LocationBlocked) {
        Write-Log 'WLAN 查询被「位置权限」挡住(Windows 11 26H2 起,netsh 读 WLAN 信息需要定位服务)' 'WARN'
        try {
            if (Repair-WlanLocationPermission) { $info = Get-WifiInfo }
        } catch {
            Write-Log ('位置权限自动修复出错: ' + $_.Exception.Message) 'WARN'
        }
    }
    Write-Log ('当前 WiFi: ' + $(if ($info.Ssid) { $info.Ssid } else { '(未连接)' }) + ' ,准备连接 ' + $Ssid)

    # 无线服务没跑就先拉起来(通常需要管理员,失败也不致命)
    try {
        $svc = Get-Service -Name 'WlanSvc' -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -ne 'Running') {
            Write-Log '无线服务 WlanSvc 未运行,尝试启动(可能需要管理员)' 'WARN'
            Start-Service -Name 'WlanSvc' -ErrorAction Stop
            Start-Sleep -Seconds 2
        }
    } catch {
        Write-Log ('启动 WlanSvc 失败:' + $_.Exception.Message) 'WARN'
    }

    # 先自动刷新一遍列表:系统缓存里的 WiFi 列表经常是过期的/空的
    $wl = Update-WifiList -Force
    if (-not $wl.TargetInProfiles -and $wl.TargetVisible -and $wl.TargetOpen) {
        Write-Log ('Windows 里没有 ' + $Ssid + ' 的配置,检测到是开放网络,自动补一份') 'WARN'
        if (Add-OpenWifiProfile -Ssid $Ssid) { $wl = Update-WifiList -Force }
    }
    if (-not $wl.TargetInProfiles -and -not $wl.LocationBlocked) {
        $visList = @($wl.Visible)
        $visTxt = $(if ($visList.Count) { ($visList | ForEach-Object { $_.Ssid }) -join ' / ' } else { '(一个都没搜到)' })
        Write-Log ('Windows 里没有保存 ' + $Ssid + ' 的配置。请先手动连接一次这个 WiFi,之后就能自动连') 'ERROR'
        Write-Log ('现在能搜到的 WiFi: ' + $visTxt) 'WARN'
        return $false
    }
    if (-not $wl.TargetVisible) {
        Write-Log ('列表里暂时看不到 ' + $Ssid + '(信号弱?),仍按已保存的配置尝试连接') 'WARN'
    }

    for ($attempt = 1; $attempt -le 2; $attempt++) {
        if ($attempt -gt 1) {
            # 刚开机/刚唤醒时无线网卡往往还没扫完,第一轮容易白等 —— 刷新列表后再试一次
            $wl2 = Update-WifiList -Force
            if (-not $wl2.TargetVisible) {
                Write-Log ('刷新列表后仍然看不到 ' + $Ssid + ',不重试了') 'WARN'
                break
            }
            Write-Log ('刷新列表后 ' + $Ssid + ' 可见(信号 ' + $wl2.TargetSignal + '%),再试一次') 'WARN'
        }

        $r = Invoke-NativeLines (Get-NetshPath) ('wlan connect name="{0}"' -f $Ssid)
        $msg = (($r | Where-Object { $_ }) -join ' ').Trim()
        if (-not $msg) { $msg = '(无输出)' }
        Write-Log ('netsh wlan connect -> ' + $msg)

        $deadline = (Get-Date).AddSeconds($TimeoutSec)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds 2
            if (Test-WifiConnected -Ssid $Ssid) {
                Write-Log ('WiFi 已连接:' + $Ssid) 'OK'
                Start-Sleep -Seconds 2   # 等 DHCP 拿到地址
                return $true
            }
        }
        Write-Log ('连接 WiFi 超时(' + $TimeoutSec + ' 秒):' + $Ssid) 'WARN'
    }

    return $false
}

function Ensure-Wifi {
    param([string]$Ssid = '')

    if ($NoWifi -or -not $script:AutoWifi) { return $true }
    if (-not $Ssid) { $Ssid = (Get-Settings).wifiSsid }
    if (Test-WifiConnected -Ssid $Ssid) { return $true }
    return (Connect-CampusWifi -Ssid $Ssid)
}

# ---------------- 网络 ----------------
function Get-NetStatus {
    # 不只看状态码,还要看内容和重定向目标:校园网常见「HTTP 被劫持到门户、甚至返回 200」
    $res = [pscustomobject]@{ Online = $false; Hijacked = $false; Detail = '所有探测都不通' }

    # 探测 1:微软的固定文本,内容对得上才算通
    try {
        $raw = & $Curl -s -w "`n%{http_code}|%{redirect_url}" --max-time 8 'http://www.msftconnecttest.com/connecttest.txt' 2>$null
        if ($raw) {
            $text = ($raw -join "`n")
            $idx = $text.LastIndexOf("`n")
            if ($idx -ge 0) {
                $body = $text.Substring(0, $idx)
                $parts = $text.Substring($idx + 1) -split '\|'
                $code = $parts[0]
                $loc = $(if ($parts.Count -gt 1) { $parts[1] } else { '' })

                if ($code -eq '200' -and $body -match 'Microsoft Connect Test') {
                    $res.Online = $true; $res.Detail = '微软探测通过'
                    return $res
                }
                if ($code -match '^30[12378]$' -and $loc -match ('ac_portal|wlanuserip|10000\.gd\.cn|' + [regex]::Escape($PortalHost))) {
                    $res.Hijacked = $true; $res.Detail = ('被重定向到门户: ' + $loc)
                    return $res
                }
                if ($code -eq '200' -and $body -match 'ac_portal|wlanuserip') {
                    $res.Hijacked = $true; $res.Detail = '直接返回了门户页面(透明劫持)'
                    return $res
                }
            }
        }
    } catch { }

    # 探测 2:要求 204 或空响应;返回 200 且带内容 = 被劫持
    try {
        $raw2 = & $Curl -s -w "`n%{http_code}" --max-time 8 'http://connect.rom.miui.com/generate_204' 2>$null
        if ($raw2) {
            $t2 = ($raw2 -join "`n")
            $i2 = $t2.LastIndexOf("`n")
            if ($i2 -ge 0) {
                $b2 = $t2.Substring(0, $i2)
                $c2 = $t2.Substring($i2 + 1)

                if ($c2 -eq '204' -or ($c2 -eq '200' -and [string]::IsNullOrWhiteSpace($b2))) {
                    $res.Online = $true; $res.Detail = 'generate_204 通过'
                    return $res
                }
                if ($c2 -eq '200') {
                    $res.Hijacked = $true; $res.Detail = '204 探测返回 200(透明劫持)'
                    return $res
                }
                $res.Detail = ('探测返回 HTTP ' + $c2)
            }
        }
    } catch { }

    return $res
}

function Test-Online {
    return (Get-NetStatus).Online
}

# ---------------- 门户配置(通用认证) ----------------
# 每家校园网的门户都不一样:接口路径、参数名、密码加密方式、成功判定各不相同。
# 所以这里不再把 Dr.COM 那套写死,而是放在 portal.json 里,按配置拼请求。
# 内置几种预设,`-Mode portal -Set <id>` 或菜单 O) 可以切换。
$script:PortalFn = Join-Path $WorkDir 'portal.json'

function Get-PortalPresets {
    $drcom = [pscustomobject]@{
        id           = 'drcom'
        name         = '中国电信 Dr.COM(RC4 + auth_tag)'
        protocol     = 'template'
        host         = '1.1.1.4'
        page         = '/ac_portal/20250303113805/pc.html?template=20250303113805&tabs=pwd&dual_stack=0&vlanid=1033'
        loginUrl     = 'http://{host}/ac_portal/login.php'
        method       = 'POST'
        contentType  = 'application/x-www-form-urlencoded'
        encrypt      = 'rc4_ms'
        template     = 'opr=pwdLogin&userName={user}&pwd={pwd}&auth_tag={tag}&rememberPwd=0'
        successMatch = '"success"\s*:\s*true'
        msgMatch     = '"msg"\s*:\s*"([^"]*)"'
        note         = '本机现用的门户。密码 = RC4(密钥为毫秒时间戳),输出小写 hex。'
    }
    $form = [pscustomobject]@{
        id           = 'form'
        name         = '通用表单(明文 / md5 / base64)'
        protocol     = 'template'
        host         = '192.168.1.1'
        page         = '/'
        loginUrl     = 'http://{host}/login'
        method       = 'POST'
        contentType  = 'application/x-www-form-urlencoded'
        encrypt      = 'none'
        template     = 'username={user}&password={pwd}'
        successMatch = ''
        msgMatch     = ''
        note         = '大多数自建/开源认证页的形态。成功与否默认靠连通性复检(成功判定留空 = 以 Test-Online 为准)。'
    }
    $srun = [pscustomobject]@{
        id           = 'srun'
        name         = '深澜 Srun(实验性)'
        protocol     = 'template'
        host         = '10.0.0.55'
        page         = '/srun_portal_pc?ac_id=1&theme=pro'
        loginUrl     = 'http://{host}/srun_portal_pc'
        method       = 'POST'
        contentType  = 'application/x-www-form-urlencoded'
        encrypt      = 'md5'
        template     = 'action=login&username={user}&password={pwd}&ac_id=1'
        successMatch = '"error"\s*:\s*"ok"'
        msgMatch     = '"error_msg"\s*:\s*"([^"]*)"'
        note         = '深澜各校版本差异大,真实密码算法常是 hmd5(带 token)。这里给的是骨架,密码规则可能还要改 encrypt 字段。'
    }
    $esurfing = [pscustomobject]@{
        id           = 'esurfing'
        name         = '广东电信 天翼校园网(enet.10000.gd.cn 客户端协议)'
        protocol     = 'esurfing'
        host         = 'enet.10000.gd.cn:10001'
        page         = ''
        loginUrl     = 'http://{host}/client/login'
        method       = 'POST'
        contentType  = 'application/json'
        encrypt      = 'none'
        template     = ''
        successMatch = '"rescode"\s*:\s*"?0"?'
        msgMatch     = '"resinfo"\s*:\s*"([^"]*)"'
        secret       = 'Eshore!@#'
        version      = '214'
        iswifi       = '4060'
        schoolid     = ''
        clientip     = ''
        nasip        = ''
        macFormat    = 'colon-lower'
        note         = '客户端协议:queryschool → advertisement.do → vchallenge → client/login,JSON + 大写 MD5 签名(密钥 Eshore!@#)。clientip/nasip 自动从门户跳转地址里取,一般不用手填。'
    }
    $jsonform = [pscustomobject]@{
        id           = 'jsonform'
        name         = '通用 JSON 表单(可带 token 两步)'
        protocol     = 'template'
        host         = '192.168.1.1'
        page         = '/'
        tokenUrl     = ''
        tokenMethod  = 'GET'
        tokenRegex   = ''
        loginUrl     = 'http://{host}/api/login'
        method       = 'POST'
        contentType  = 'application/json; charset=UTF-8'
        encrypt      = 'none'
        template     = '{"username":"{user}","password":"{pwd}"}'
        successMatch = '"success"\s*:\s*true|"code"\s*:\s*"?0"?'
        msgMatch     = '"message"\s*:\s*"([^"]*)"'
        note         = '内置预设都套不上时用这个:JSON 体 + 可选 token 两步(填 tokenUrl/tokenRegex 就先去取 token)。成功判定留空则靠连通性。'
    }
    $srun2 = [pscustomobject]@{
        id           = 'srun2'
        name         = '深澜 Srun v2(token + hmd5,骨架)'
        protocol     = 'template'
        host         = '10.0.0.55'
        page         = '/srun_portal_pc?ac_id=1&theme=pro'
        tokenUrl     = '/cgi-bin/get_challenge?callback=jsonp&username={user}&ip={clientip}'
        tokenMethod  = 'GET'
        tokenRegex   = '"token"\s*:\s*"([^"]+)"'
        loginUrl     = 'http://{host}/cgi-bin/srun_portal'
        method       = 'GET'
        contentType  = 'application/x-www-form-urlencoded'
        encrypt      = 'md5'
        template     = 'callback=jsonp&action=login&username={user}&password={pwd}&ac_id=1&ip={clientip}&chksum={token}&info=&n=200&type=1&os=Windows&name=Windows&double_stack=0'
        successMatch = '"res"\s*:\s*"ok"'
        msgMatch     = '"error_msg"\s*:\s*"([^"]*)"'
        note         = '骨架:token 两步已经接好,但真实密码算法各校不同(常见 hmd5 = md5(md5(pwd)+token),chksum 也要用 token 拼)。按抓包改 encrypt/template。'
    }
    $ruijie = [pscustomobject]@{
        id           = 'ruijie'
        name         = '锐捷 ePortal(骨架)'
        protocol     = 'template'
        host         = '192.168.0.1'
        page         = '/eportal/portal/login.jsp'
        tokenUrl     = ''
        tokenMethod  = 'GET'
        tokenRegex   = ''
        loginUrl     = 'http://{host}/eportal/portal/login'
        method       = 'POST'
        contentType  = 'application/x-www-form-urlencoded'
        encrypt      = 'none'
        template     = 'userId={user}&password={pwd}&service=&queryString=wlanuserip%3D{clientip}&operatorPwd=&operatorUserId=&validcode=&passwordEncrypt=false'
        successMatch = '"result"\s*:\s*"success"'
        msgMatch     = '"message"\s*:\s*"([^"]*)"'
        note         = '骨架:参数名各版本不同(userId/userName、password/pwd)。queryString 一般要从门户页里抠出完整 wlanuserip/wlanacname,可以填 tokenUrl/tokenRegex 先取。'
    }
    return @($drcom, $esurfing, $form, $jsonform, $srun, $srun2, $ruijie)
}

function Get-PortalPreset {
    param([string]$Id)
    foreach ($p in (Get-PortalPresets)) { if ($p.id -eq $Id) { return $p } }
    return $null
}

function Save-PortalProfile {
    param($Profile)
    try {
        # 只写这个协议用得上的字段,别把别的预设的 null 塞进文件
        $keep = @('id', 'name', 'protocol', 'host', 'page', 'loginUrl', 'method', 'contentType', 'encrypt',
            'template', 'successMatch', 'msgMatch', 'secret', 'version', 'iswifi', 'schoolid', 'clientip', 'nasip',
            'macFormat', 'tokenUrl', 'tokenMethod', 'tokenRegex', 'tokenData', 'referer', 'headers', 'note')
        $obj = [ordered]@{}
        foreach ($k in $keep) {
            $v = $Profile.PSObject.Properties[$k]
            if ($v -and $null -ne $v.Value) { $obj[$k] = $v.Value }
        }
        ($obj | ConvertTo-Json) | Set-Content -Path $script:PortalFn -Encoding UTF8
        return $true
    } catch {
        Write-Log ('保存 portal.json 失败:' + $_.Exception.Message) 'ERROR'
        return $false
    }
}

function Get-PortalProfile {
    # 读 portal.json;缺字段用 Dr.COM 默认补齐(老版本升级过来的情况)
    $def = Get-PortalPreset 'drcom'
    if (-not (Test-Path $script:PortalFn)) {
        [void](Save-PortalProfile $def)
        return $def
    }
    try {
        $p = Get-Content $script:PortalFn -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in @('id', 'name', 'protocol', 'host', 'page', 'loginUrl', 'method', 'contentType', 'encrypt', 'template', 'successMatch', 'msgMatch', 'secret', 'version', 'iswifi', 'schoolid', 'clientip', 'nasip', 'macFormat', 'tokenUrl', 'tokenMethod', 'tokenRegex', 'tokenData', 'referer', 'headers', 'note')) {
            if (-not $p.PSObject.Properties[$k]) { $p | Add-Member -NotePropertyName $k -NotePropertyValue $def.$k -Force }
        }
        if (-not $p.host) { $p.host = $def.host }
        if (-not $p.loginUrl) { $p.loginUrl = $def.loginUrl }
        if (-not $p.template -and [string]$p.protocol -ne 'esurfing') { $p.template = $def.template }
        if (-not $p.encrypt) { $p.encrypt = 'none' }
        return $p
    } catch {
        Write-Log ('portal.json 解析失败,按默认 Dr.COM 处理:' + $_.Exception.Message) 'WARN'
        return $def
    }
}

function Set-PortalProfile {
    param([string]$Id)
    $preset = Get-PortalPreset $Id
    if (-not $preset) { Write-Log ('没有这个预设: ' + $Id) 'WARN'; return $false }
    if (-not (Save-PortalProfile $preset)) { return $false }
    Write-Log ('门户配置已切换为: ' + $preset.name) 'OK'
    return $true
}

function Get-StringHash {
    param([string]$Text, [string]$Algo)
    $sha = [System.Security.Cryptography.HashAlgorithm]::Create($Algo)
    $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text))
    return (($bytes | ForEach-Object { $_.ToString('x2') }) -join '')
}

function Get-PortalEncodedPassword {
    param([string]$Plain, [string]$Mode, [string]$Tag)
    switch ($Mode) {
        'rc4_ms' { return (Get-Rc4Hex -Text $Plain -Key $Tag) }
        'md5' { return (Get-StringHash -Text $Plain -Algo 'MD5') }
        'sha1' { return (Get-StringHash -Text $Plain -Algo 'SHA1') }
        'base64' { return [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Plain)) }
        default { return $Plain }
    }
}

function Get-JsonText {
    # 塞进 JSON 字符串里用(转义但不带引号)
    param([string]$Text)
    $s = ConvertTo-EsJsonString $Text
    return $s.Substring(1, $s.Length - 2)
}

function Get-PortalClientIp {
    # 门户眼里的「客户端 IP」:先看 302 跳转里的 wlanuserip,取不到就用出口网卡
    $cap = Get-CaptiveRedirect
    $p = Get-EsParam $cap
    if ($p.ClientIp) { return $p.ClientIp }
    return (Get-EgressNic).Ip
}

function Invoke-PortalLogin {
    # 按 portal.json 拼请求;返回 @{ Success; Msg; Raw }
    #
    # 支持的门户形态:
    #   1) 一次表单 POST/GET            —— template + encrypt
    #   2) 先取 token 再登录             —— tokenUrl + tokenRegex(很多 Srun / 锐捷 / 自建门户要)
    #   3) JSON 体                      —— contentType 里带 json 就行,占位符按 JSON 转义
    #   4) 多步客户端协议               —— protocol=esurfing(单独实现)
    param([string]$Account, [string]$PlainPassword)

    $pf = Get-PortalProfile
    if ([string]$pf.protocol -eq 'esurfing') {
        return (Invoke-ESurfingLogin -Account $Account -PlainPassword $PlainPassword)
    }
    $h = $(if ($PortalHost) { $PortalHost } else { $pf.host })
    $ct = $(if ($pf.contentType) { [string]$pf.contentType } else { 'application/x-www-form-urlencoded' })
    $isJson = [bool]($ct -match 'json')

    if ($pf.page) {
        try {
            & $Curl -s -b $CookieJar -c $CookieJar -o NUL --max-time 15 (('http://{0}{1}') -f $h, $pf.page) 2>$null | Out-Null
        } catch {
            Write-Log ("访问门户页失败(继续尝试登录):$($_.Exception.Message)") 'WARN'
        }
    }

    $tag = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds().ToString()
    $encPwd = Get-PortalEncodedPassword -Plain $PlainPassword -Mode $pf.encrypt -Tag $tag

    # 客户端 IP:有门户要求带上,不要求就别去探(省时间)
    $clientIp = ''
    $needIp = ([string]$pf.template -match '\{clientip\}') -or ([string]$pf.tokenUrl -match '\{clientip\}') -or
              ([string]$pf.loginUrl -match '\{clientip\}') -or ([string]$pf.tokenUrl -match '\{ip\}')
    if ($needIp) { $clientIp = Get-PortalClientIp }

    # ---- 第一步:取 token ----
    $token = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$pf.tokenUrl)) {
        $tu = [string]$pf.tokenUrl
        $tu = $tu.Replace('{host}', $h).Replace('{clientip}', $clientIp).Replace('{ip}', $clientIp)
        $tu = $tu.Replace('{user}', [uri]::EscapeDataString($Account)).Replace('{tag}', $tag)
        if ($tu -notmatch '^https?://') { $tu = 'http://' + $h + $tu }

        $td = [string]$pf.tokenData
        $td = $td.Replace('{host}', $h).Replace('{clientip}', $clientIp).Replace('{ip}', $clientIp)
        $td = $td.Replace('{user}', [uri]::EscapeDataString($Account)).Replace('{tag}', $tag)

        $tokRaw = @()
        try {
            if ([string]$pf.tokenMethod -match '^(?i)post$') {
                $tokRaw = & $Curl -s -b $CookieJar -c $CookieJar -A $UA -X POST $tu `
                    -H ('Content-Type: ' + $ct) --data $td --max-time 15 2>$null
            } else {
                $tokRaw = & $Curl -s -b $CookieJar -c $CookieJar -A $UA --max-time 15 $tu 2>$null
            }
        } catch {
            Write-Log ('取 token 失败(继续试登录):' + $_.Exception.Message) 'WARN'
        }
        $tokText = ($tokRaw -join "`n")
        if ([string]$pf.tokenRegex) {
            if ($tokText -match [string]$pf.tokenRegex) { $token = $Matches[1] }
        }
        if ($token) { Write-Log ('取到 token:' + $token) }
        else { Write-Log ('没取到 token(检查 tokenUrl / tokenRegex)。返回:' + $tokText.Substring(0, [Math]::Min(160, $tokText.Length))) 'WARN' }
    }

    # ---- 第二步:拼登录请求 ----
    $uEsc = $(if ($isJson) { Get-JsonText $Account } else { [uri]::EscapeDataString($Account) })
    $pEsc = $(if ($isJson) { Get-JsonText $encPwd } else { [uri]::EscapeDataString($encPwd) })
    $plEsc = $(if ($isJson) { Get-JsonText $PlainPassword } else { [uri]::EscapeDataString($PlainPassword) })

    $body = [string]$pf.template
    $body = $body.Replace('{user}', $uEsc)
    $body = $body.Replace('{plain}', $plEsc)
    $body = $body.Replace('{pwd}', $pEsc)
    $body = $body.Replace('{tag}', $tag)
    $body = $body.Replace('{host}', $h)
    $body = $body.Replace('{token}', $(if ($isJson) { Get-JsonText $token } else { [uri]::EscapeDataString($token) }))
    $body = $body.Replace('{clientip}', $clientIp)

    $url = ([string]$pf.loginUrl).Replace('{host}', $h).Replace('{clientip}', $clientIp)

    $hdrArgs = @()
    if ($pf.referer) {
        $hdrArgs += @('-H', ('Referer: ' + ([string]$pf.referer).Replace('{host}', $h)))
    }
    if ($pf.headers) {
        # 多条用 ;; 隔开(避免和 Cookie 里的 ; 打架)
        foreach ($one in ([string]$pf.headers -split ';;')) {
            $one = $one.Trim()
            if ($one) { $hdrArgs += @('-H', $one.Replace('{host}', $h)) }
        }
    }

    # 请求体/响应都走临时文件:JSON 里的引号在 PS 5.1 直接传参会被搞坏,响应落盘再按 UTF-8 读也不会乱码
    $id = [guid]::NewGuid().ToString('N')
    $bodyFn = Join-Path $env:TEMP ('campusnet-login-' + $id + '.dat')
    $outFn = Join-Path $env:TEMP ('campusnet-login-' + $id + '.out')
    $httpCode = ''
    try {
        [System.IO.File]::WriteAllText($bodyFn, $body, (New-Object System.Text.UTF8Encoding($false)))
        $argList = @('-s', '-b', $CookieJar, '-c', $CookieJar, '-A', $UA, '-o', $outFn, '-w', '%{http_code}', '--max-time', '20')
        if ($hdrArgs.Count -gt 0) { $argList += $hdrArgs }

        if ([string]$pf.method -match '^(?i)get$') {
            $sep = $(if ($url -match '\?') { '&' } else { '?' })
            $httpCode = ((& $Curl @argList ($url + $sep + $body) 2>$null) -join '').Trim()
        } else {
            $argList += @('-X', 'POST', '-H', ('Content-Type: ' + $ct), '--data-binary', ('@' + $bodyFn), $url)
            $httpCode = ((& $Curl @argList 2>$null) -join '').Trim()
        }

        $text = ''
        if (Test-Path -LiteralPath $outFn) {
            $text = [System.IO.File]::ReadAllText($outFn, [System.Text.Encoding]::UTF8)
        }
    } finally {
        Remove-Item -LiteralPath $bodyFn, $outFn -Force -ErrorAction SilentlyContinue
    }

    if (-not $text) { throw ('登录接口没有返回任何内容(HTTP ' + $httpCode + ';检查是否还连着校园网、门户地址对不对)') }

    $ok = $false
    if ($pf.successMatch) {
        $ok = [bool]($text -match $pf.successMatch)
    } else {
        # 没配成功特征 = 以连通性为准
        $ok = Test-Online
    }

    $msg = ''
    if ($pf.msgMatch) {
        if ($text -match $pf.msgMatch) { $msg = $Matches[1] }
    } elseif ($text -match '"msg"\s*:\s*"([^"]*)"') {
        $msg = $Matches[1]
    }
    if (-not $msg) { $msg = $text.Substring(0, [Math]::Min(120, $text.Length)) }

    return [pscustomobject]@{ Success = $ok; Msg = $msg; Raw = $text; HttpCode = $httpCode }
}

# ---------------- 天翼校园网(广东电信 enet.10000.gd.cn) ----------------
# 天翼不是「一张表单 POST 完事」,而是客户端协议,四步走:
#   1) client/queryschool   学校 ID(只用来取会话 Cookie,拿不到也能登)
#   2) advertisement.do     顺手薅一个会话 Cookie
#   3) client/vchallenge    取 challenge(服务端下发的短信码,不用 OCR —— 这正是它比网页版好使的地方)
#   4) client/login         真正登录
# 每个请求都是 JSON,都带 timestamp 和大写 MD5 签名,密钥是固定串 Eshore!@#。
# clientip / nasip 不用手填:未认证时访问任意网站会被 302 到
#   http://enet.10000.gd.cn:10001/?wlanuserip=<clientip>&wlanacip=<nasip>
# 从跳转地址里抠出来即可。

function ConvertTo-EsJsonString {
    param([string]$Text)
    $s = [string]$Text
    $s = $s.Replace('\', '\\').Replace('"', '\"')
    return '"' + $s + '"'
}

function New-EsSign {
    param([string]$Buffer)
    return ((Get-StringHash -Text $Buffer -Algo 'MD5')).ToUpperInvariant()
}

function Format-MacAddress {
    param([string]$Mac, [string]$Style = 'colon-lower')
    $hex = ([string]$Mac -replace '[^0-9A-Fa-f]', '').ToLowerInvariant()
    if ($hex.Length -ne 12) { return [string]$Mac }
    $pairs = @()
    for ($i = 0; $i -lt 12; $i += 2) { $pairs += $hex.Substring($i, 2) }
    $joined = ($pairs -join ':')
    switch ($Style) {
        'colon-upper' { return $joined.ToUpperInvariant() }
        'plain-lower' { return $hex }
        'plain-upper' { return $hex.ToUpperInvariant() }
        default { return $joined }
    }
}

function Get-LocalNetInfo {
    # 按 IP 反查网卡,拿到 MAC(签名里要用)
    param([string]$Ip)
    $r = [pscustomobject]@{ Ip = [string]$Ip; Mac = ''; IfIndex = 0; Alias = '' }
    if ([string]::IsNullOrWhiteSpace($Ip)) { return $r }
    try {
        $a = Get-NetIPAddress -IPAddress $Ip -ErrorAction Stop | Select-Object -First 1
        if ($a) {
            $r.IfIndex = [int]$a.InterfaceIndex
            $r.Alias = [string]$a.InterfaceAlias
            $na = Get-NetAdapter -InterfaceIndex $r.IfIndex -ErrorAction SilentlyContinue
            if ($na -and $na.MacAddress) { $r.Mac = ([string]$na.MacAddress).ToLowerInvariant() }
        }
    } catch { }
    if (-not $r.Mac) {
        try {
            $c = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ErrorAction Stop |
                Where-Object { $_.IPAddress -contains $Ip } | Select-Object -First 1
            if ($c -and $c.MACAddress) { $r.Mac = ([string]$c.MACAddress).ToLowerInvariant() }
        } catch { }
    }
    return $r
}

# ------- 自动识别门户(连上校园网后浏览器会自动弹认证页,本质就是「外网请求被 302」) -------
$script:CaptiveProbes = @(
    'http://www.msftconnecttest.com/connecttest.txt',
    'http://connect.rom.miui.com/generate_204',
    'http://www.qq.com/',
    'http://189.cn/'
)

function Get-CaptiveRedirect {
    # 复制一次「浏览器弹认证页」:探一个外网地址,把 302 的 Location 抓下来 —— 那门户地址
    param([string[]]$ProbeUrl = @())
    $probes = $(if ($ProbeUrl -and $ProbeUrl.Count -gt 0) { $ProbeUrl } else { $script:CaptiveProbes })
    foreach ($t in $probes) {
        try {
            $probeHost = ''
            try { $probeHost = ([uri]$t).Host } catch { }
            $line = ((& $Curl -s -o NUL -A $UA -w '%{http_code}|%{redirect_url}' --max-time 6 $t 2>$null) -join '').Trim()
            if (-not $line) { continue }
            $parts = $line -split '\|', 2
            $code = $parts[0]
            $loc = $(if ($parts.Count -gt 1) { [string]$parts[1] } else { '' })
            if (-not $loc -or $code -notmatch '^30[12378]$') { continue }

            $locHost = ''
            try { $locHost = ([uri]$loc).Host } catch { }

            # 带认证参数 = 门户;换了主机且不是"正常跳转" = 门户
            $isPortal = ($loc -match 'wlanuserip=|wlanacip=|ac_portal|srun|eportal|/portal/|:8080/')
            if (-not $isPortal -and $locHost) {
                $isPortal = ($locHost -ne $probeHost) -and
                            ($locHost -notmatch '(^|\.)(qq\.com|bing\.com|baidu\.com|msftconnecttest\.com|msftncsi\.com|gstatic\.com|apple\.com|miui\.com)$')
            }
            if ($isPortal) { return $loc }
        } catch { }
    }
    return ''
}

function Get-HttpsPortalHint {
    # 浏览器「自动弹认证页」很多是被 HTTPS 劫持:请求被换成了门户页面。
    # 这里用 -k 忽略证书错误取回页面,只要页面已经不是原来那个站点,就把标题和页面里的域名捞出来。
    param([string]$Probe = 'https://www.baidu.com/')
    $id = [guid]::NewGuid().ToString('N')
    $outFn = Join-Path $env:TEMP ('campusnet-https-' + $id + '.html')
    try {
        $code = ((& $Curl -sk -L -A $UA --max-time 8 -o $outFn -w '%{http_code}' $Probe 2>$null) -join '').Trim()
        if (-not $code -or $code -eq '000') { return $null }
        if (-not (Test-Path -LiteralPath $outFn)) { return $null }
        $html = [System.IO.File]::ReadAllText($outFn, [System.Text.Encoding]::UTF8)
        if (-not $html) { return $null }
        # 真页面里一定有 baidu.com;没有就是被换掉了
        if ($html -match 'baidu\.com') { return $null }

        $title = ''
        if ($html -match '<title>([^<]{1,80})</title>') { $title = $Matches[1].Trim() }
        $hosts = @()
        foreach ($m in [regex]::Matches($html, 'https?://([a-zA-Z0-9\.\-]+(?::\d+)?)')) {
            $hv = $m.Groups[1].Value.ToLowerInvariant()
            if ($hv -match 'w3\.org|jquery|cdnjs|unpkg|jsdelivr|bdstatic|bcebos|gstatic') { continue }
            if ($hosts -notcontains $hv) { $hosts += $hv }
        }
        return [pscustomobject]@{
            Title = $(if ($title) { $title } else { '(没有标题)' })
            Hosts = @($hosts | Select-Object -First 6)
            Code  = $code
            Size  = $html.Length
        }
    } catch { return $null } finally {
        Remove-Item -LiteralPath $outFn -Force -ErrorAction SilentlyContinue
    }
}

function Get-CarrierHint {
    # 运营商线索(只看域名和 SSID)。注意:运营商只决定出口线路,认证系统是学校另外买的
    param([string]$Url = '', [string]$Ssid = '')
    $t = (([string]$Url) + ' ' + ([string]$Ssid)).ToLowerInvariant()
    if ($t -match 'esurfing|10000\.gd|189\.cn|chinatelecom|chinanet|-ct\b|电信') { return '中国电信(天翼)' }
    if ($t -match 'cmcc|10086|hejiaoyuan|chinamobile|移动') { return '中国移动(CMCC-EDU / 和校园)' }
    if ($t -match 'chinaunicom|10010|unicom|(^|\W)wo(\W|$)|联通') { return '中国联通(沃)' }
    return ''
}

function Get-PortalFingerprint {
    # 光凭门户地址就能认出是哪家,顺带把「主机 + 路径」拿来当配置
    param([string]$Url)
    $f = [pscustomobject]@{ Vendor = '未知'; Id = ''; Host = ''; Page = '/'; Url = [string]$Url; LoginUrl = '' }
    if ([string]::IsNullOrWhiteSpace($Url)) { return $f }
    try {
        $u = [uri]$Url
        $f.Host = [string]$u.Authority
        $f.Page = [string]$u.PathAndQuery
        if (-not $f.Page) { $f.Page = '/' }
    } catch { return $f }

    # 顺序有讲究:天翼的跳转也带 wlanuserip,和锐捷/深澜撞车,所以先认域名
    if ($Url -match '10000\.gd\.cn|10000\.cn|\besurfing\b|enet\.|:10001(?:/|\?|$)') {
        $f.Vendor = '广东电信 天翼校园网'
        $f.Id = 'esurfing'
        $f.Host = 'enet.10000.gd.cn:10001'
        $f.Page = '/'
        $f.LoginUrl = 'http://{host}/client/login'
        return $f
    }
    if ($Url -match 'srun|深澜|get_challenge|cgi-bin/srun') {
        $f.Vendor = '深澜 Srun'
        $f.Id = 'srun2'
        $f.LoginUrl = 'http://{host}/cgi-bin/srun_portal'
        return $f
    }
    if ($Url -match 'ac_portal|drcom|1\.1\.1\.4|城市热点') {
        $f.Vendor = '中国电信 Dr.COM'
        $f.Id = 'drcom'
        $f.LoginUrl = 'http://{host}/ac_portal/login.php'
        return $f
    }
    if ($Url -match 'eportal|锐捷|ruijie|wlan_user_ip') {
        $f.Vendor = '锐捷 ePortal(RG)'
        $f.Id = 'ruijie'
        $f.LoginUrl = 'http://{host}/eportal/portal/login'
        return $f
    }
    if ($Url -match 'sangfor|深信服') { $f.Vendor = '深信服 Sangfor'; return $f }
    if ($Url -match 'h3c|huawei|hwportal|portal\.cgi') { $f.Vendor = 'H3C / 华为'; return $f }
    if ($Url -match '10\.\d+\.\d+\.\d+|192\.168\.\d+\.\d+|172\.(?:1[6-9]|2\d|3[01])\.') {
        $f.Vendor = '内网门户(私有地址,看不出厂商)'
        return $f
    }
    return $f
}

function Invoke-PortalAdapt {
    # 看现在被劫持到哪,顺势把 portal.json 改成对得上的那套
    param([string]$Url = '', [switch]$Quiet)

    $pf = Get-PortalProfile
    $capUrl = $(if ($Url) { $Url } else { Get-CaptiveRedirect })
    if (-not $capUrl) {
        if (-not $Quiet) { Write-Log '没抓到门户跳转(可能已经在线,或者这个网络不劫持 HTTP)' 'WARN' }
        return [pscustomobject]@{ Changed = $false; Vendor = ''; Url = ''; Reason = '没抓到跳转' }
    }

    $fp = Get-PortalFingerprint $capUrl
    if (-not $Quiet) { Write-Log ('抓到门户: ' + $capUrl) }
    if (-not $Quiet) { Write-Log ('识别为: ' + $fp.Vendor + $(if ($fp.Id) { ' (' + $fp.Id + ')' } else { '(没有现成预设)' })) }

    $curId = [string]$pf.id
    $reasons = @()
    $changed = $false

    # 只对「实测过」的预设做自动切换;骨架预设(深澜/锐捷)只提示,别替用户瞎改配置
    $autoOk = @('drcom', 'esurfing', 'form')

    if ($fp.Id -and $fp.Id -ne $curId) {
        if ($autoOk -contains $fp.Id) {
            $preset = Get-PortalPreset $fp.Id
            if ($preset) {
                $pf = $preset
                $changed = $true
                $reasons += ('协议 ' + $curId + ' → ' + $fp.Id)
            }
        } else {
            Write-Log ('识别到「' + $fp.Vendor + '」,但内置的 ' + $fp.Id + ' 只是骨架,没自动切。要试就先 -Mode portal -PortalId ' + $fp.Id + ' 再按抓包补字段') 'WARN'
            return [pscustomobject]@{ Changed = $false; Vendor = $fp.Vendor; Url = $capUrl; Reason = ('需要手工适配 ' + $fp.Id) }
        }
    }

    if ($fp.Id -and $fp.Id -eq [string]$pf.id -and $autoOk -contains [string]$pf.id) {
        if ($fp.Host -and $fp.Host -ne [string]$pf.host) {
            $pf.host = $fp.Host; $changed = $true; $reasons += ('门户地址 ' + $fp.Host)
        }
        # 页面路径只对「页面式」门户有意义(天翼是固定 API,不用改)
        if ([string]$pf.id -ne 'esurfing' -and $fp.Page -and $fp.Page -ne '/' -and $fp.Page -ne [string]$pf.page) {
            $pf.page = $fp.Page; $changed = $true; $reasons += ('认证页 ' + $fp.Page)
        }
    }

    if ($changed) {
        if (Save-PortalProfile $pf) {
            Write-Log ('门户配置已自动更新: ' + ($reasons -join '、')) 'OK'
        } else {
            $changed = $false
            $reasons += '写入 portal.json 失败'
        }
    } else {
        if (-not $Quiet) { Write-Log '门户配置和当前环境一致,没动' }
    }

    return [pscustomobject]@{ Changed = $changed; Vendor = $fp.Vendor; Url = $capUrl; Reason = ($reasons -join '、') }
}

function Get-EsParam {
    param([string]$Url)
    $o = [pscustomobject]@{ Url = [string]$Url; ClientIp = ''; NasIp = '' }
    if ([string]::IsNullOrWhiteSpace($Url)) { return $o }
    if ($Url -match '[?&]wlanuserip=([^&]+)') { $o.ClientIp = [uri]::UnescapeDataString($Matches[1]) }
    if ($Url -match '[?&]wlanacip=([^&]+)') { $o.NasIp = [uri]::UnescapeDataString($Matches[1]) }
    return $o
}

function Invoke-EsJsonPost {
    # curl 输出先落盘再按 UTF-8 读,免得中文响应被控制台代码页搞成乱码
    param([string]$Url, [string]$Json, [string]$Cookie = '')
    $id = [guid]::NewGuid().ToString('N')
    $bodyFn = Join-Path $env:TEMP ('campusnet-es-' + $id + '.json')
    $outFn = Join-Path $env:TEMP ('campusnet-es-' + $id + '.out')
    try {
        [System.IO.File]::WriteAllText($bodyFn, $Json, (New-Object System.Text.UTF8Encoding($false)))
        $argList = @('-s', '--max-time', '15', '-A', $UA,
            '-H', 'Content-Type: application/json; charset=UTF-8',
            '--data-binary', ('@' + $bodyFn), '-o', $outFn, '-w', '%{http_code}')
        if ($Cookie) { $argList += @('-H', ('Cookie: ' + $Cookie)) }
        $argList += $Url
        $httpCode = ((& $Curl @argList 2>$null) -join '').Trim()
        $text = ''
        if (Test-Path -LiteralPath $outFn) {
            $text = [System.IO.File]::ReadAllText($outFn, [System.Text.Encoding]::UTF8).Trim()
        }
        return [pscustomobject]@{ Code = $httpCode; Text = $text }
    } finally {
        Remove-Item -LiteralPath $bodyFn, $outFn -Force -ErrorAction SilentlyContinue
    }
}

function Get-EsCookie {
    param([string]$Base, [string]$SchoolId)
    if (-not $SchoolId) { return '' }
    try {
        $hdr = & $Curl -s -D - -o NUL -A $UA --max-time 8 ($Base + '/advertisement.do?' + $SchoolId) 2>$null
        $c = @()
        foreach ($line in ($hdr -split "`r?`n")) {
            if ($line -match '^Set-Cookie:\s*([^;]+)') { $c += $Matches[1].Trim() }
        }
        return ($c -join '; ')
    } catch { return '' }
}

function Get-EgressNic {
    # 默认路由出口网卡(802.1X / NAT / 抓不到 clientip 时全靠它)
    try {
        $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
            Sort-Object { [int]$_.RouteMetric } | Select-Object -First 1
        if ($route) {
            $na = Get-NetAdapter -InterfaceIndex $route.InterfaceIndex -ErrorAction SilentlyContinue
            $ip = Get-NetIPAddress -InterfaceIndex $route.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.IPAddress -notmatch '^169\.254\.' } | Select-Object -First 1
            return [pscustomobject]@{
                IfIndex = [int]$route.InterfaceIndex
                Mac     = $(if ($na -and $na.MacAddress) { ([string]$na.MacAddress).ToLowerInvariant() } else { '' })
                Alias   = $(if ($na) { [string]$na.Name } else { '' })
                Ip      = $(if ($ip) { [string]$ip.IPAddress } else { '' })
            }
        }
    } catch { }
    return [pscustomobject]@{ IfIndex = 0; Mac = ''; Alias = ''; Ip = '' }
}

function Get-DefaultGatewayIp {
    # 默认网关:门户域名打不开时,很多学校的认证页就架在网关上(10.x.0.1 这种)
    try {
        $r = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
            Sort-Object { [int]$_.RouteMetric } | Select-Object -First 1
        if ($r -and $r.NextHop -match '^\d+\.\d+\.\d+\.\d+$' -and $r.NextHop -ne '0.0.0.0') {
            return [string]$r.NextHop
        }
    } catch { }
    return ''
}

function Resolve-EsContext {
    param([string]$ClientIp = '', [string]$NasIp = '')
    # 只有缺参数时才去撞 302(每次登录都撞一遍太慢)
    $capUrl = ''
    if (-not $ClientIp -or -not $NasIp) { $capUrl = Get-CaptiveRedirect }
    $p = Get-EsParam $capUrl
    if (-not $ClientIp) { $ClientIp = $p.ClientIp }
    if (-not $NasIp) { $NasIp = $p.NasIp }

    $egress = $null
    if (-not $ClientIp) {
        # 没被劫持(可能已在线):用默认路由出口网卡的 IPv4 碰碰运气
        $egress = Get-EgressNic
        if ($egress.Ip) { $ClientIp = $egress.Ip }
    }

    # MAC 优先按 clientip 找;clientip 不是本机地址(经过 NAT / 前面还有路由器)时退回出口网卡
    $mac = ''
    if ($ClientIp) { $mac = (Get-LocalNetInfo -Ip $ClientIp).Mac }
    if (-not $mac) {
        if (-not $egress) { $egress = Get-EgressNic }
        $mac = $egress.Mac
    }
    return [pscustomobject]@{ ClientIp = $ClientIp; NasIp = $NasIp; Mac = $mac; Redirect = $capUrl }
}

function Get-EsMessage {
    param([string]$Code, [string]$Info = '')
    $map = @{
        '0'        = '认证成功'
        '13002000' = '已经在线(重复登录)'
        '13012000' = '密码错误'
        '13004000' = '用户认证失败(多半是账号/密码不对)'
        '11064000' = '登录太频繁,过几分钟再试'
        '13018000' = '服务器禁止网页/第三方认证,只能改用官方客户端'
        '13016000' = '这个账号没有订购校园网产品'
        '11063000' = 'challenge 对不上(重试一次通常就好)'
        '13005000' = '认证超时(等几秒重试)'
        '13000000' = '参数被拒(检查 clientip/nasip/MAC 是否对得上本机)'
    }
    if ($Code -and $map.ContainsKey($Code)) { return ($map[$Code] + '(rescode ' + $Code + ')') }
    if ($Info) { return ($Info + $(if ($Code) { '(rescode ' + $Code + ')' } else { '' })) }
    return ('未知结果' + $(if ($Code) { '(rescode ' + $Code + ')' } else { '' }))
}

function Invoke-ESurfingLogin {
    param([string]$Account, [string]$PlainPassword)

    $pf = Get-PortalProfile
    $api = [string]$pf.host
    if ([string]::IsNullOrWhiteSpace($api)) { $api = 'enet.10000.gd.cn:10001' }
    if ($api -notmatch ':') { $api = $api + ':10001' }
    $secret = [string]$pf.secret; if (-not $secret) { $secret = 'Eshore!@#' }
    $version = [string]$pf.version; if (-not $version) { $version = '214' }
    $iswifi = [string]$pf.iswifi; if (-not $iswifi) { $iswifi = '4060' }
    $macStyle = [string]$pf.macFormat; if (-not $macStyle) { $macStyle = 'colon-lower' }
    $base = 'http://' + $api

    $ctx = Resolve-EsContext -ClientIp ([string]$pf.clientip) -NasIp ([string]$pf.nasip)
    $clientIp = [string]$ctx.ClientIp
    $nasIp = [string]$ctx.NasIp
    $mac = Format-MacAddress -Mac $ctx.Mac -Style $macStyle

    Write-Log ('天翼认证参数:clientip=' + $(if ($clientIp) { $clientIp } else { '?' }) + ' nasip=' + $(if ($nasIp) { $nasIp } else { '?' }) + ' mac=' + $(if ($mac) { $mac } else { '?' }))
    if ($ctx.Redirect) { Write-Log ('门户跳转:' + $ctx.Redirect) }

    if (-not $clientIp) { throw '取不到 clientip(当前可能已在线,或者不在校园网)。先用浏览器点一次认证页,或运行 -Mode portal -DetectPortal 看现场' }
    if (-not $mac) { throw '取不到本机网卡 MAC(签名要用)。确认 clientip 就是本机 IP' }
    if (-not $nasIp) { Write-Log '拿不到 nasip(接入服务器 IP),继续尝试;签名可能对不上' 'WARN' }

    # 1) 学校 ID
    $schoolId = [string]$pf.schoolid
    if (-not $schoolId) {
        $ts = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds().ToString()
        $sign = New-EsSign -Buffer ($clientIp + $nasIp + $mac + $ts + $secret)
        $json = '{"clientip":' + (ConvertTo-EsJsonString $clientIp) + ',"nasip":' + (ConvertTo-EsJsonString $nasIp) +
                ',"mac":' + (ConvertTo-EsJsonString $mac) + ',"timestamp":' + $ts + ',"authenticator":' + (ConvertTo-EsJsonString $sign) + '}'
        try {
            $r = Invoke-EsJsonPost -Url ($base + '/client/queryschool') -Json $json
            if ($r.Text -match '"schoolid"\s*:\s*"?(\d+)"?') { $schoolId = $Matches[1] }
            Write-Log ('queryschool:' + $(if ($schoolId) { '学校 ID ' + $schoolId } else { '没取到(不致命)' }))
        } catch {
            Write-Log ('queryschool 失败(不致命):' + $_.Exception.Message) 'WARN'
        }
    }

    # 2) 会话 Cookie
    $cookie = ''
    if ($schoolId) {
        try { $cookie = Get-EsCookie -Base $base -SchoolId $schoolId } catch { }
        if ($cookie) { Write-Log ('拿到会话 Cookie:' + $cookie) }
    }

    $lastText = ''
    $lastMsg = '没有发起请求'
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        # 3) vchallenge
        $ts = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds().ToString()
        $sign = New-EsSign -Buffer ($version + $clientIp + $nasIp + $mac + $ts + $secret)
        $json = '{"version":' + (ConvertTo-EsJsonString $version) + ',"username":' + (ConvertTo-EsJsonString $Account) +
                ',"clientip":' + (ConvertTo-EsJsonString $clientIp) + ',"nasip":' + (ConvertTo-EsJsonString $nasIp) +
                ',"mac":' + (ConvertTo-EsJsonString $mac) + ',"timestamp":' + (ConvertTo-EsJsonString $ts) +
                ',"authenticator":' + (ConvertTo-EsJsonString $sign) + '}'

        $ch = Invoke-EsJsonPost -Url ($base + '/client/vchallenge') -Json $json -Cookie $cookie
        $challenge = ''
        if ($ch.Text -match '"challenge"\s*:\s*"([^"]*)"') { $challenge = $Matches[1] }
        if (-not $challenge) {
            $rc = ''
            if ($ch.Text -match '"rescode"\s*:\s*"?(-?\d+)"?') { $rc = $Matches[1] }
            $lastText = $ch.Text
            Write-Log ('vchallenge 没给 challenge(HTTP ' + $ch.Code + '):' + $ch.Text) 'WARN'
            return [pscustomobject]@{ Success = $false; Msg = ('取不到 challenge ' + (Get-EsMessage -Code $rc)); Raw = $ch.Text }
        }
        Write-Log ('challenge:' + $challenge)

        # 4) client/login
        $ts = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds().ToString()
        $sign = New-EsSign -Buffer ($clientIp + $nasIp + $mac + $ts + $challenge + $secret)
        $json = '{"username":' + (ConvertTo-EsJsonString $Account) + ',"password":' + (ConvertTo-EsJsonString $PlainPassword) +
                ',"clientip":' + (ConvertTo-EsJsonString $clientIp) + ',"nasip":' + (ConvertTo-EsJsonString $nasIp) +
                ',"mac":' + (ConvertTo-EsJsonString $mac) + ',"iswifi":' + (ConvertTo-EsJsonString $iswifi) +
                ',"timestamp":' + (ConvertTo-EsJsonString $ts) + ',"authenticator":' + (ConvertTo-EsJsonString $sign) +
                ',"verificationcode":' + (ConvertTo-EsJsonString $challenge) + '}'

        $lr = Invoke-EsJsonPost -Url ($base + '/client/login') -Json $json -Cookie $cookie
        $lastText = $lr.Text
        $rc = ''
        if ($lr.Text -match '"rescode"\s*:\s*"?(-?\d+)"?') { $rc = $Matches[1] }
        $info = ''
        if ($lr.Text -match '"resinfo"\s*:\s*"([^"]*)"') { $info = $Matches[1] }
        $lastMsg = Get-EsMessage -Code $rc -Info $info
        Write-Log ('client/login 返回(HTTP ' + $lr.Code + '):' + $lr.Text)

        if ($rc -eq '0') {
            if (Test-Online) { Write-Log '认证成功,连通性复检也通过' 'OK' }
            else { Write-Log '认证成功,但立刻复检还没通(有些环境要几秒才放行)' 'WARN' }
            return [pscustomobject]@{ Success = $true; Msg = $lastMsg; Raw = $lr.Text }
        }

        if ((@('11063000', '13005000') -contains $rc) -and $attempt -lt 3) {
            Write-Log ('第 ' + $attempt + ' 次没成功(' + $lastMsg + '),重拿 challenge 再试') 'WARN'
            Start-Sleep -Milliseconds 900
            continue
        }
        return [pscustomobject]@{ Success = $false; Msg = $lastMsg; Raw = $lr.Text }
    }

    return [pscustomobject]@{ Success = $false; Msg = $lastMsg; Raw = $lastText }
}

function Invoke-EsDetect {
    # 现场勘查:把天翼环境的关键参数和服务器原始返回都打出来,好对症改 portal.json
    $pf = Get-PortalProfile
    $api = [string]$pf.host; if ([string]::IsNullOrWhiteSpace($api)) { $api = 'enet.10000.gd.cn:10001' }
    if ($api -notmatch ':') { $api = $api + ':10001' }
    $base = 'http://' + $api
    $secret = [string]$pf.secret; if (-not $secret) { $secret = 'Eshore!@#' }
    $version = [string]$pf.version; if (-not $version) { $version = '214' }
    $macStyle = [string]$pf.macFormat; if (-not $macStyle) { $macStyle = 'colon-lower' }

    Write-Stdout ''
    Write-Stdout '=== 天翼校园网 现场勘查 ==='
    $ctx = Resolve-EsContext
    Write-Stdout ('  重定向地址: ' + $(if ($ctx.Redirect) { $ctx.Redirect } else { '(没有 302 —— 可能已经在线,或者不在校园网)' }))
    Write-Stdout ('  clientip:   ' + $(if ($ctx.ClientIp) { $ctx.ClientIp } else { '?' }))
    Write-Stdout ('  nasip:      ' + $(if ($ctx.NasIp) { $ctx.NasIp } else { '?' }))
    $mac = Format-MacAddress -Mac $ctx.Mac -Style $macStyle
    Write-Stdout ('  mac:        ' + $(if ($mac) { $mac } else { '?' }) + '   (格式 ' + $macStyle + ')')
    if (-not $ctx.ClientIp) { Write-Stdout '  提示: 拿不到参数就先把浏览器打开认证页再跑一次'; return }

    $nic = Get-LocalNetInfo -Ip $ctx.ClientIp
    Write-Stdout ('  网卡:       ' + $nic.Alias + '  ' + $nic.Mac)

    $ts = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds().ToString()
    $sign = New-EsSign -Buffer ($ctx.ClientIp + $ctx.NasIp + $mac + $ts + $secret)
    $json = '{"clientip":' + (ConvertTo-EsJsonString $ctx.ClientIp) + ',"nasip":' + (ConvertTo-EsJsonString $ctx.NasIp) +
            ',"mac":' + (ConvertTo-EsJsonString $mac) + ',"timestamp":' + $ts + ',"authenticator":' + (ConvertTo-EsJsonString $sign) + '}'
    $sch = Invoke-EsJsonPost -Url ($base + '/client/queryschool') -Json $json
    Write-Stdout ('  queryschool: HTTP ' + $sch.Code + '  ' + $sch.Text)

    $ts = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds().ToString()
    $sign = New-EsSign -Buffer ($version + $ctx.ClientIp + $ctx.NasIp + $mac + $ts + $secret)
    $json = '{"version":' + (ConvertTo-EsJsonString $version) + ',"username":"","clientip":' + (ConvertTo-EsJsonString $ctx.ClientIp) +
            ',"nasip":' + (ConvertTo-EsJsonString $ctx.NasIp) + ',"mac":' + (ConvertTo-EsJsonString $mac) +
            ',"timestamp":' + (ConvertTo-EsJsonString $ts) + ',"authenticator":' + (ConvertTo-EsJsonString $sign) + '}'
    $ch = Invoke-EsJsonPost -Url ($base + '/client/vchallenge') -Json $json
    Write-Stdout ('  vchallenge:  HTTP ' + $ch.Code + '  ' + $ch.Text)
    Write-Stdout ''
    Write-Stdout '  这两步能返回 rescode 0 / challenge,说明协议对得上,直接登录即可:'
    Write-Stdout '    .\Connect-CampusNet.ps1 -Mode portal -PortalId esurfing'
    Write-Stdout '    .\Connect-CampusNet.ps1 -Mode login'
    Write-Stdout '  如果 rescode 不是 0,把上面这几行发给作者(里面不含密码)'
    Write-Stdout ''
}

function Show-PortalProfile {
    $pf = Get-PortalProfile
    $preset = Get-PortalPreset $pf.id
    Write-Stdout ''
    Write-Stdout '=== 当前门户配置 ==='
    Write-Stdout ('  名称:   ' + $pf.name + $(if (-not $preset) { '  (自定义)' } else { '' }))
    Write-Stdout ('  协议:   ' + $(if ([string]$pf.protocol -eq 'esurfing') { 'esurfing(天翼客户端协议,多步 JSON)' } else { 'template(单次表单提交)' }))
    Write-Stdout ('  门户:   ' + $pf.host + $(if ($pf.page) { $pf.page } else { '' }))
    Write-Stdout ('  登录:   ' + $pf.method + ' ' + $pf.loginUrl)
    if ([string]$pf.protocol -eq 'esurfing') {
        Write-Stdout ('  版本:   version=' + $pf.version + '   iswifi=' + $pf.iswifi + '   MAC 格式=' + $pf.macFormat)
        Write-Stdout ('  学校:   ' + $(if ($pf.schoolid) { 'ID ' + $pf.schoolid + '(固定)' } else { '自动查询' }))
        Write-Stdout ('  签名:   MD5(大写),密钥 ' + $pf.secret + '   —— 串接顺序照抄客户端协议')
        if ($pf.clientip -or $pf.nasip) { Write-Stdout ('  手填:   clientip=' + $pf.clientip + ' nasip=' + $pf.nasip) }
        else { Write-Stdout '  参数:   clientip/nasip 自动从门户跳转地址里取' }
    } else {
        Write-Stdout ('  加密:   ' + $pf.encrypt + '   (rc4_ms / md5 / sha1 / base64 / none)')
        if ($pf.tokenUrl) {
            Write-Stdout ('  取 token: ' + $pf.tokenMethod + ' ' + $pf.tokenUrl)
            Write-Stdout ('  正则:     ' + $pf.tokenRegex)
        }
        Write-Stdout ('  参数:   ' + $pf.template)
        if ($pf.headers) { Write-Stdout ('  额外头: ' + $pf.headers) }
        Write-Stdout ('  判定:   ' + $(if ($pf.successMatch) { '正则 ' + $pf.successMatch } else { '靠连通性复检' }))
    }
    if ($pf.note) { Write-Stdout ('  说明:   ' + $pf.note) }
    Write-Stdout ('  文件:   ' + $script:PortalFn)
    Write-Stdout ''
    Write-Stdout '  内置预设:'
    foreach ($p in (Get-PortalPresets)) { Write-Stdout ('    ' + $p.id.PadRight(10) + $p.name) }
    Write-Stdout ''
}

function Invoke-PortalDetect {
    # 抓门户页,按特征猜是哪家(猜不准就算了,不强行改配置)
    $pf = Get-PortalProfile
    $ssid = ''
    try { $ssid = (Get-WifiInfo).Ssid } catch { }

    # 先看「被劫持到哪了」—— 天翼的跳转地址里直接带着 clientip/nasip
    $capUrl = Get-CaptiveRedirect
    if ($capUrl) {
        Write-Stdout ('未认证,被重定向到: ' + $capUrl)
    } else {
        if (Test-Online) { Write-Stdout '当前已经在线,门户不会被劫持(要测识别:断开 WiFi 再连一次)' }
        else { Write-Stdout '没抓到 302 —— 这个网络可能不在 HTTP 层拦(只拦 HTTPS),或者门户直接挂在网关上' }

        # 门户域名不通时,先看网关:很多学校(尤其运营商宿舍网)的认证页就架在网关上
        $gw0 = Get-DefaultGatewayIp
        if ($gw0) {
            foreach ($gu in @(('http://' + $gw0 + '/'), ('http://' + $gw0 + ':8080/'))) {
                $gc = ''
                try { $gc = ((& $Curl -s -o NUL -A $UA -w '%{http_code}' --max-time 5 $gu 2>$null) -join '').Trim() } catch { }
                Write-Stdout ('  试网关: ' + $gu + ' → HTTP ' + $gc)
                if ($gc -match '^[23]') {
                    $gHtml = ''
                    try { $gHtml = (& $Curl -s -L -A $UA --max-time 8 $gu 2>$null) -join "`n" } catch { }
                    $gt = ''
                    if ($gHtml -match '<title>([^<]{1,80})</title>') { $gt = $Matches[1] }
                    Write-Stdout ('    网关上有页面: ' + $(if ($gt) { $gt } else { '(没有 title)' }) + '  ' + $gHtml.Length + ' 字节')
                    if ($gHtml -match 'portal|login|认证|eportal|srun|wlanuserip|username') {
                        Write-Stdout '    页面里有认证特征 → 这就是门户,把上面这个网关地址填进 host'
                    }
                }
            }
        }
    }

    $fp = Get-PortalFingerprint $capUrl
    $carrier = Get-CarrierHint -Url $capUrl -Ssid $ssid
    if ($fp.Id -or $carrier) {
        Write-Stdout ('识别结果: ' + $fp.Vendor + $(if ($fp.Id) { '(' + $fp.Id + ')' } else { '' }))
        Write-Stdout ('运营商线索: ' + $(if ($carrier) { $carrier } else { '(看不出,或门户在内网私有地址)' }))
        Write-Stdout '提示: 运营商(电信/移动/联通)只决定出口线路,认证系统是学校另买的 —— 同一所学校换运营商,门户可能一模一样。'
    }
    if ($capUrl -and $fp.Id -eq 'esurfing') {
        Write-Stdout ''
        Invoke-EsDetect
        return
    }

    $h = $(if ($PortalHost) { $PortalHost } else { $pf.host })
    $page = $(if ($pf.page) { $pf.page } else { '/' })
    $url = ('http://{0}{1}' -f $h, $page)

    # HTTP 层没抓到跳转时,再看 HTTPS 层:浏览器"自动弹认证页"多半就是把 HTTPS 请求劫持了
    if (-not $capUrl) {
        $httpsHint = Get-HttpsPortalHint
        if ($httpsHint) {
            Write-Stdout ''
            Write-Stdout ('HTTPS 层抓到劫持页: ' + $httpsHint.Title)
            if ($httpsHint.Hosts.Count -gt 0) {
                Write-Stdout ('  页面上出现的域名: ' + ($httpsHint.Hosts -join ' , '))
                Write-Stdout '  → 门户地址多半在其中;用 -Mode portal -AutoPortal 试,或把域名填进 host'
            }
        }
    }

    Write-Stdout ('探测: ' + $url)
    $html = ''
    try {
        $html = (& $Curl -s -L -A $UA --max-time 15 $url 2>$null) -join "`n"
    } catch { }
    if (-not $html) {
        if (-not $capUrl) { Write-Stdout '连门户页都打不开(地址错了,或者当前不在校园网)' }
        else { Write-Stdout '门户页打不开(可能是 HTTPS 门户,或者要先带着跳转参数进去)' }
        return
    }

    $vendor = '未知'
    $probe = $html + ' ' + $page
    if ($probe -match 'srun_portal|get_challenge|cgi-bin/srun|深澜') { $vendor = '深澜 Srun' }
    elseif ($probe -match 'ac_portal|Dr\.COM|drcom') { $vendor = '中国电信 Dr.COM' }
    elseif ($probe -match 'eportal|锐捷|ruijie') { $vendor = '锐捷 ePortal(RG)' }
    elseif ($probe -match 'sangfor|深信服') { $vendor = '深信服 Sangfor' }
    elseif ($probe -match 'h3c|华为|huawei|hwportal') { $vendor = 'H3C / 华为' }
    elseif ($probe -match 'wlanuserip') { $vendor = '通用 Portal(带 wlanuserip 参数)' }

    Write-Stdout ('页面识别: ' + $vendor)
    if ($probe -match 'ac_portal') { Write-Stdout '  出现 ac_portal 字样 → 和 drcom 预设兼容' }
    if ($probe -match 'get_challenge|srun_portal') { Write-Stdout '  出现 srun 特征 → 先切 srun2(骨架),token 两步已接好,密码算法还要对一下' }
    if ($probe -match 'eportal') { Write-Stdout '  出现 eportal 字样 → 锐捷,切 ruijie(骨架)后按抓包补参数名' }
    Write-Stdout ('  页面长度: ' + $html.Length + ' 字节')
}

# ---------------- 时间校正 ----------------
function Get-PortalServerTime {
    # 未认证时门户也能访问,正好用它响应里的 Date 头在登录前校准本机时间
    $hdr = & $Curl -s -D - -o NUL --max-time 8 ('http://{0}{1}' -f $PortalHost, $PortalPage) 2>$null
    if (-not $hdr) { return $null }
    foreach ($line in ($hdr -split "`r?`n")) {
        if ($line -match '^Date:\s*(.+)$') {
            try { return [DateTimeOffset]::Parse($Matches[1].Trim(), [Globalization.CultureInfo]::InvariantCulture) }
            catch { return $null }
        }
    }
    return $null
}

function Test-Elevated {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]::new($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Sync-Clock {
    param([switch]$Always)

    $srv = Get-PortalServerTime
    if (-not $srv) {
        Write-Log '取不到门户服务器时间(可能没连校园网),跳过时间校正' 'WARN'
        return $false
    }

    $skew = ($srv.UtcDateTime - [DateTime]::UtcNow).TotalSeconds
    $msg = ('本机时间与门户服务器相差 {0:N1} 秒' -f $skew)

    if ([Math]::Abs($skew) -le 1) {
        Write-Log ($msg + ',正常')
        return $true
    }

    if ([Math]::Abs($skew) -le 3 -and -not $Always) {
        Write-Log ($msg + ',在可接受范围内')
        return $true
    }

    if (Test-Elevated) {
        try {
            Set-Date -Adjust ([TimeSpan]::FromSeconds([Math]::Round($skew))) | Out-Null
            Write-Log ($msg + ' → 已自动校正') 'OK'
            return $true
        } catch {
            Write-Log ('校正时间失败:' + $_.Exception.Message) 'ERROR'
            return $false
        }
    }

    # 非管理员:试试 w32tm(校园网常屏蔽 NTP,失败很正常)
    try {
        & w32tm /resync /force 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { Write-Log ($msg + ' → 已通过 w32tm 校正') 'OK'; return $true }
    } catch { }

    Write-Log ($msg + ' —— 偏差大时登录会被拒。请用管理员运行本脚本选「校正系统时间」,或打开 设置 > 时间和语言 > 自动设置时间') 'WARN'
    return $false
}

# ---------------- 多账号存储 ----------------
function Get-Store {
    if (-not (Test-Path $AccountsFn)) {
        if (Test-Path $LegacyFn) {
            try {
                $old = Get-Content $LegacyFn -Raw -Encoding UTF8 | ConvertFrom-Json
                $store = [pscustomobject]@{
                    active   = $old.user
                    accounts = @([pscustomobject]@{ user = $old.user; scope = $old.scope; secret = $old.secret; label = '' })
                }
                Save-Store -Store $store
                Remove-Item $LegacyFn -Force -ErrorAction SilentlyContinue
                Write-Log ('已把旧的 cred.json 迁移为 accounts.json(账号:' + $old.user + ')') 'OK'
                return $store
            } catch {
                Write-Log ('迁移旧凭据失败:' + $_.Exception.Message) 'WARN'
            }
        }
        return [pscustomobject]@{ active = ''; accounts = @() }
    }

    try {
        $s = Get-Content $AccountsFn -Raw -Encoding UTF8 | ConvertFrom-Json
        $s.accounts = @($s.accounts | Where-Object { $_ })
        if (-not $s.active -and $s.accounts.Count -gt 0) { $s.active = $s.accounts[0].user }
        return $s
    } catch {
        Write-Log ('读取 accounts.json 失败:' + $_.Exception.Message) 'ERROR'
        return [pscustomobject]@{ active = ''; accounts = @() }
    }
}

function Save-Store {
    param($Store)
    $out = [pscustomobject]@{
        active   = $Store.active
        accounts = @($Store.accounts | ForEach-Object {
                [pscustomobject]@{ user = $_.user; scope = $_.scope; secret = $_.secret; label = $_.label }
            })
    }
    $out | ConvertTo-Json -Depth 5 | Set-Content -Path $AccountsFn -Encoding UTF8
}

function Protect-Secret {
    param([string]$PlainPassword, [switch]$UseMachineScope)

    if ($UseMachineScope) {
        Add-Type -AssemblyName System.Security
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainPassword)
        $enc = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
        return @{ scope = 'machine'; secret = [Convert]::ToBase64String($enc) }
    }
    $sec = ConvertTo-SecureString $PlainPassword -AsPlainText -Force
    return @{ scope = 'user'; secret = (ConvertFrom-SecureString $sec) }
}

function Unprotect-Secret {
    param($Entry)
    if ($Entry.scope -eq 'machine') {
        Add-Type -AssemblyName System.Security
        $enc = [Convert]::FromBase64String($Entry.secret)
        $bytes = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $enc, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    }
    $sec = ConvertTo-SecureString $Entry.secret
    return [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
}

function Get-AccountList {
    $store = Get-Store
    return @($store.accounts)
}

function Find-Account {
    param([string]$Account)
    $store = Get-Store
    if (-not $Account) { $Account = $store.active }
    if (-not $Account) { return $null }
    return @($store.accounts | Where-Object { $_.user -eq $Account })[0]
}

function Get-ActiveCred {
    $store = Get-Store
    $entry = Find-Account -Account $store.active
    if (-not $entry) { return $null }
    try {
        return [pscustomobject]@{ User = $entry.user; Password = (Unprotect-Secret -Entry $entry); Label = $entry.label }
    } catch {
        Write-Log ('解密账号 ' + $entry.user + ' 失败(可能换了电脑/Windows 用户):' + $_.Exception.Message) 'ERROR'
        return $null
    }
}

function Set-Account {
    param([string]$Account, [string]$PlainPassword, [string]$Note, [switch]$MakeActive, [switch]$UseMachineScope)

    $store = Get-Store
    $list = @($store.accounts | Where-Object { $_.user -ne $Account })
    $sec = Protect-Secret -PlainPassword $PlainPassword -UseMachineScope:$UseMachineScope
    $list += [pscustomobject]@{ user = $Account; scope = $sec.scope; secret = $sec.secret; label = $Note }
    $store.accounts = $list
    if ($MakeActive -or -not $store.active) { $store.active = $Account }
    Save-Store -Store $store
    Write-Log ('账号已保存:' + $Account + $(if ($Note) { '(' + $Note + ')' } else { '' }) + '  加密范围:' + $sec.scope) 'OK'
}

function Remove-Account {
    param([string]$Account)
    $store = Get-Store
    $store.accounts = @($store.accounts | Where-Object { $_.user -ne $Account })
    if ($store.active -eq $Account) {
        if ($store.accounts.Count -gt 0) { $store.active = $store.accounts[0].user } else { $store.active = '' }
    }
    Save-Store -Store $store
    Write-Log ('已删除账号:' + $Account) 'OK'
}

function Set-ActiveAccount {
    param([string]$Account)
    $store = Get-Store
    if (-not (@($store.accounts | Where-Object { $_.user -eq $Account }).Count)) {
        Write-Log ('没有这个账号:' + $Account) 'ERROR'
        return $false
    }
    $store.active = $Account
    Save-Store -Store $store
    Write-Log ('当前账号已切换为:' + $Account) 'OK'
    return $true
}

# ---------------- 状态文件(供托盘读取) ----------------
function Write-State {
    param([bool]$Online, [string]$Account, [string]$Note = '')

    try {
        $obj = [pscustomobject]@{
            online    = $Online
            account   = $Account
            note      = $Note
            pid       = $PID
            timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
        $obj | ConvertTo-Json -Compress | Set-Content -Path $StateFn -Encoding UTF8
    } catch { }
}

function Read-State {
    if (-not (Test-Path $StateFn)) { return $null }
    try { return (Get-Content $StateFn -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

function Get-StateAge {
    param($State)
    try {
        $t = [datetime]::ParseExact($State.timestamp, 'yyyy-MM-dd HH:mm:ss', $null)
        return ((Get-Date) - $t).TotalSeconds
    } catch { return 999999 }
}

# ---------------- 登录一次 ----------------
function Invoke-LoginOnce {
    param($Cred, [switch]$NoAdapt)

    try {
        $res = Invoke-PortalLogin -Account $Cred.User -PlainPassword $Cred.Password
    } catch {
        Write-Log ('登录请求失败:' + $_.Exception.Message) 'ERROR'
        $res = $null
    }

    if ($res -and $res.Success) {
        Write-Log ('登录成功: ' + $res.Msg) 'OK'
        if (Test-Online) { Write-Log '连通性复检通过,可以上网了' 'OK' }
        else { Write-Log '返回成功但复检未通过(有些环境要几秒才放行)' 'WARN' }
        return $true
    }

    if ($res) { Write-Log ('登录失败: ' + $res.Msg) 'ERROR' }

    # 失败后先别急着报错:看一眼门户是不是变了(换校区/门户改版都会这样),变了就换配置再试一次。
    # 只在"确实被劫持"时才去抓门户 —— 网络整个不通时去探没意义,只会让守护多等几十秒
    if (-not $NoAdapt) {
        $st = Get-NetStatus
        if ($st.Hijacked) {
            $a = Invoke-PortalAdapt
            if ($a.Changed) {
                Write-Log ('门户没匹配上,已自动切到「' + $a.Vendor + '」(' + $a.Reason + '),重试一次') 'OK'
                return (Invoke-LoginOnce -Cred $Cred -NoAdapt)
            }
        } else {
            Write-Log ('网络状态:' + $st.Detail + ',不去抓门户了')
        }
    }
    return $false
}

# ---------------- 交互输入 ----------------
function Read-PlainPassword {
    param([string]$Prompt = '请输入密码')
    $sec = Read-Host $Prompt -AsSecureString
    return [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
}

function Resolve-Cred {
    param([string]$Account, [string]$PlainPassword, [switch]$AllowPrompt)

    if ($Account) {
        if ($PlainPassword) { return [pscustomobject]@{ User = $Account; Password = $PlainPassword } }
        $entry = Find-Account -Account $Account
        if ($entry) {
            try { return [pscustomobject]@{ User = $entry.user; Password = (Unprotect-Secret -Entry $entry) } } catch { }
        }
    }

    $cred = Get-ActiveCred
    if ($cred) { return $cred }

    if (-not $AllowPrompt) { return $null }

    if (-not $Account) { $Account = Read-Host '请输入校园网账号(一般是电信手机号)' }
    if (-not $PlainPassword) { $PlainPassword = Read-PlainPassword }
    return [pscustomobject]@{ User = $Account; Password = $PlainPassword }
}

# ================= 各模式 =================
function Invoke-Setup {
    param([switch]$IsChange)

    Write-Stdout ''
    if ($IsChange) { Write-Stdout '=== 添加账号 / 更换密码 ===' } else { Write-Stdout '=== 添加账号 ===' }

    $store = Get-Store
    if ($store.accounts.Count -gt 0) {
        Write-Stdout '已有账号:'
        foreach ($a in $store.accounts) {
            $mark = $(if ($a.user -eq $store.active) { '*' } else { ' ' })
            Write-Stdout ('   ' + $mark + ' ' + $a.user + $(if ($a.label) { '  (' + $a.label + ')' } else { '' }))
        }
    }

    $acct = Read-Host '请输入账号(一般回车取消;同名账号=改密码)'
    if ([string]::IsNullOrWhiteSpace($acct)) { return }
    $acct = $acct.Trim()

    $note = Read-Host '备注(可留空,如 主号 / 备用)'
    $pwd = Read-PlainPassword
    if ([string]::IsNullOrWhiteSpace($pwd)) { Write-Log '密码不能为空' 'ERROR'; return }

    Set-Account -Account $acct -PlainPassword $pwd -Note $note -MakeActive -UseMachineScope:$MachineScope

    Write-Stdout ''
    Write-Stdout '验证一下这组账号密码(会真的发一次登录请求)…'
    $ok = Invoke-LoginOnce -Cred ([pscustomobject]@{ User = $acct; Password = $pwd })
    Write-Stdout ''
    if ($ok) { Write-Stdout '验证通过。' }
    else { Write-Stdout '若显示「用户名或密码错误」请重新添加;若显示「已在线」之类,凭据大概率没问题。' }
}

function Invoke-Switch {
    $store = Get-Store
    if ($store.accounts.Count -eq 0) { Write-Stdout '还没有任何账号,先添加'; return }
    if ($store.accounts.Count -eq 1) { Write-Stdout '只保存了一个账号,不用切换'; return }

    Write-Stdout ''
    Write-Stdout '=== 选择当前使用的账号 ==='
    $i = 0
    foreach ($a in $store.accounts) {
        $i++
        $mark = $(if ($a.user -eq $store.active) { '*' } else { ' ' })
        Write-Stdout ('  ' + $i + ') ' + $mark + ' ' + $a.user + $(if ($a.label) { '  (' + $a.label + ')' } else { '' }))
    }
    $sel = Read-Host '输入序号(回车取消)'
    if ([string]::IsNullOrWhiteSpace($sel)) { return }
    $idx = 0
    if (-not [int]::TryParse($sel, [ref]$idx) -or $idx -lt 1 -or $idx -gt $store.accounts.Count) {
        Write-Stdout '序号无效'; return
    }
    [void](Set-ActiveAccount -Account $store.accounts[$idx - 1].user)
    Write-Stdout '提示:如果守护正在运行,切换后它会用新账号重连(或重启一下守护)'
}

function Invoke-Delete {
    $store = Get-Store
    if ($store.accounts.Count -eq 0) { Write-Stdout '还没有任何账号'; return }

    Write-Stdout ''
    Write-Stdout '=== 删除账号 ==='
    $i = 0
    foreach ($a in $store.accounts) {
        $i++
        Write-Stdout ('  ' + $i + ') ' + $a.user + $(if ($a.label) { '  (' + $a.label + ')' } else { '' }))
    }
    $sel = Read-Host '输入要删除的序号(回车取消)'
    if ([string]::IsNullOrWhiteSpace($sel)) { return }
    $idx = 0
    if (-not [int]::TryParse($sel, [ref]$idx) -or $idx -lt 1 -or $idx -gt $store.accounts.Count) {
        Write-Stdout '序号无效'; return
    }
    Remove-Account -Account $store.accounts[$idx - 1].user
}

function Show-Accounts {
    $store = Get-Store
    Write-Stdout ''
    if ($store.accounts.Count -eq 0) { Write-Stdout '还没有保存任何账号'; return }
    Write-Stdout ('已保存 ' + $store.accounts.Count + ' 个账号:')
    foreach ($a in $store.accounts) {
        $mark = $(if ($a.user -eq $store.active) { '   <= 当前使用' } else { '' })
        Write-Stdout ('  ' + $a.user + $(if ($a.label) { '  (' + $a.label + ')' } else { '' }) + $mark)
    }
}

# ================= 守护 =================
# ---------------- 休眠 / 唤醒后重新认证 ----------------
function Invoke-ResumeReauth {
    param($Cred, [int]$GapSec = 0)

    Write-Log ('检测到系统从休眠/睡眠恢复(睡了 ' + $GapSec + ' 秒),重新确认网络…') 'WARN'
    try {
        $s = Get-Settings
        $s.lastResume = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Save-Settings -Settings $s
    } catch { }

    # 刚醒过来无线服务/网卡可能还没就绪:最多 3 轮,每轮等 3 秒
    for ($i = 1; $i -le 3; $i++) {
        Start-Sleep -Seconds 3
        if (Test-Online) {
            Write-Log ('唤醒后网络正常(第 ' + $i + ' 次检查)') 'OK'
            if ($i -gt 1) {
                [void](Send-Notify -Title '校园网' -Message '唤醒后网络已自行恢复' -Tag 'resumeok' -MinIntervalSec 120)
            }
            return $true
        }
        if (-not $NoWifi) { [void](Ensure-Wifi -Ssid $WifiSsid) }
        if ($Cred -and (Invoke-LoginOnce -Cred $Cred)) {
            Write-Log ('唤醒后已重新认证(第 ' + $i + ' 轮)') 'OK'
            [void](Send-Notify -Title '校园网已连接' -Message '唤醒后已自动重新认证' `
                    -Tag 'resumeok' -MinIntervalSec 120 -Actions @('打开菜单|campusnet:menu'))
            return $true
        }
    }

    Write-Log '唤醒后重认证失败,转入自动修复' 'ERROR'
    [void](Send-Notify -Title '校园网重认证失败' -Message '唤醒后连不上,已开始自动修复' -Kind 'warn' `
            -Tag 'resumefail' -MinIntervalSec 120 -Actions @('一键修复|campusnet:repair', '查看日志|campusnet:log'))
    return $false
}

function Invoke-Watch {
    $mutex = New-Object System.Threading.Mutex($false, 'Local\CampusNetAutoLogin')
    if (-not $mutex.WaitOne(0)) {
        Write-Log '已有守护进程在运行,本实例退出' 'WARN'
        return
    }

    $cred = Resolve-Cred -Account $User -Password $Password -AllowPrompt:(-not $Background)
    if (-not $cred) {
        Write-Log '没有可用的账号密码。后台模式不会弹输入框,请先手动运行为账户设置' 'ERROR'
        Write-State -Online $false -Account '' -Note 'no-credential'
        return
    }

    Write-Log ('守护启动:账号 ' + $cred.User + ',每 ' + $IntervalSec + ' 秒检测一次,pid=' + $PID)

    $wifiInfo = Get-WifiInfo
    $wifiNow = $(if ($wifiInfo.Ssid) { $wifiInfo.Ssid } else { '(未连接)' })
    if ($NoWifi -or -not $script:AutoWifi) {
        Write-Log ('WiFi 当前: ' + $wifiNow + '(WiFi 自动连接已关闭)')
    } else {
        Write-Log ('WiFi 当前: ' + $wifiNow + ' ,目标: ' + $WifiSsid)
    }

    [void](Sync-Clock)

    $wasOnline = $null
    $failCount = 0
    $cycle = 0

    while ($true) {
        try {
            $ns = Get-NetStatus
            $online = $ns.Online
            if ($online) {
                if ($wasOnline -eq $false) { Write-Log '网络已恢复' 'OK' }
                $wasOnline = $true
                $failCount = 0
                Write-State -Online $true -Account $cred.User
            } else {
                if ($wasOnline -ne $false) {
                    if ($ns.Hijacked) { Write-Log ('检测到未认证/被劫持(' + $ns.Detail + '),开始重连…') 'WARN' }
                    else { Write-Log '检测到掉线,开始重连…' 'WARN' }
                }
                $wasOnline = $false
                if (-not $NoWifi) {
                    # 每轮都重新读设置,这样在托盘/菜单里改了开关能立刻生效,不用重启守护
                    $cs = Get-Settings
                    $script:AutoWifi = [bool]$cs.autoWifi
                    [void](Ensure-Wifi -Ssid $cs.wifiSsid)
                }
                Write-State -Online $false -Account $cred.User -Note 'logging-in'

                if (Invoke-LoginOnce -Cred $cred) {
                    $wasOnline = $true
                    $failCount = 0
                    Write-State -Online $true -Account $cred.User
                } else {
                    $failCount++
                    Write-State -Online $false -Account $cred.User -Note 'login-failed'
                    if ($failCount -eq 3) {
                        Write-Log '连续 3 次失败:可能改密码了或门户改版。请打开菜单更新凭据' 'ERROR'
                        [void](Send-Notify -Title '校园网登录连续失败' -Message '可能改密码了,或门户改版。点下面打开菜单更新凭据' -Kind 'warn' `
                                -Tag 'loginfail' -MinIntervalSec 1800 -Actions @('打开菜单|campusnet:menu', '查看日志|campusnet:log'))
                    }

                    # 掉线不是密码问题时,常见原因就那三个:DNS 缓存、旧 Cookie、网卡僵住
                    if ($failCount -eq 2 -or $failCount -eq 4) {
                        $lvl = $(if ($failCount -ge 4) { 'full' } else { 'safe' })
                        try {
                            $rep = Invoke-Repair -Level $lvl -Auto -Why ('连续 ' + $failCount + ' 次登录失败')
                            if ($rep.Ran) {
                                $doneNames = @($rep.Steps | Where-Object { $_.Ok } | ForEach-Object { $_.Name })
                                $failNames = @($rep.Steps | Where-Object { -not $_.Ok } | ForEach-Object { $_.Name })
                                $txt = $(if ($doneNames.Count) { $doneNames -join '、' } else { '没有可做的' })
                                if ($failNames.Count) { $txt += ';没成功:' + ($failNames -join '、') }
                                if ($rep.Online) { $txt += ' → 已恢复在线' } else { $txt += ' → 仍不通,继续重试' }
                                [void](Send-Notify -Title $(if ($rep.Online) { '自动修复成功' } else { '自动修复完成' }) -Message $txt `
                                        -Kind $(if ($rep.Online) { 'info' } else { 'warn' }) -Tag 'autorepair' -MinIntervalSec 120)
                            }
                        } catch { Write-Log ('自动修复异常:' + $_.Exception.Message) 'ERROR' }
                    }
                }
            }

            $cycle++
            # 每 24 小时做一次门户改版自检(只抓页面和脚本,不发登录请求)
            $due = $true
            $lastChk = (Get-Settings).lastPortalCheck
            if ($lastChk) {
                try { $due = ((Get-Date) - [datetime]::ParseExact($lastChk, 'yyyy-MM-dd HH:mm:ss', $null)).TotalHours -ge 24 }
                catch { $due = $true }
            }
            if ($due) { [void](Invoke-PortalCheck) }

            if ($cycle % 60 -eq 0) { [void](Sync-Clock) }
        } catch {
            Write-Log ('守护循环异常:' + $_.Exception.Message) 'ERROR'
        }

        # 睡眠/休眠会让这一觉远超检测间隔 —— 睡醒了就立刻重新认证,不等下一轮
        $sleepStart = Get-Date
        Start-Sleep -Seconds $IntervalSec
        $sleptFor = ((Get-Date) - $sleepStart).TotalSeconds
        if ($sleptFor -gt ($IntervalSec + 15)) {
            try { [void](Invoke-ResumeReauth -Cred $cred -GapSec ([int]$sleptFor)) }
            catch { Write-Log ('唤醒后重新认证异常:' + $_.Exception.Message) 'ERROR' }
        }
    }
}

# ================= 通知:统一走 Windows 通知中心 =================
# 托盘气泡(ShowBalloonTip)在 Win10/11 上经常被系统吞掉,而且留不在通知中心里。
# 这里改用 WinRT 的 Toast:有标题/正文/按钮,点按钮能直接唤起本程序;
# 并且在「设置 > 通知」里会出现「校园网自动登录」这一条独立开关 ——
# 前提是有一个带 AppUserModelID 的开始菜单快捷方式,下面会自动建(存在就跳过)。
$script:ToastAumid   = 'CampusNet.AutoLogin'
$script:ToastAppName = '校园网自动登录'
$script:ToastState   = $null      # $null=还没探测过 $true=通知中心可用 $false=只能退回气泡
$script:ToastSeen    = @{}        # 去重:tag -> 上次发送时间
$script:UriVbs       = Join-Path $WorkDir 'campusnet-uri.vbs'
$script:AppIconFile  = Join-Path $WorkDir 'campusnet.ico'

# 写带 AUMID 的快捷方式:必须自己用 IShellLink 建(先用 WScript.Shell 建好再改,
# 属性存储是只读的,会报 STG_E_ACCESSDENIED)
$script:LinkHelperCs = @'
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace CampusNetShell {
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PKey { public Guid fmtid; public int pid; }

    [StructLayout(LayoutKind.Explicit)]
    public struct PVal { [FieldOffset(0)] public short vt; [FieldOffset(8)] public IntPtr p; }

    [ComImport, Guid("000214F9-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellLinkW {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string name);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder dir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string dir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder args, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string args);
        void GetHotkey(out short hotkey);
        void SetHotkey(short hotkey);
        void GetShowCmd(out int showCmd);
        void SetShowCmd(int showCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder iconPath, int cch, out int iconIndex);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string iconPath, int iconIndex);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pathRel, uint reserved);
        void Resolve(IntPtr hwnd, uint flags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string path);
    }

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropStore {
        void GetCount(out uint c);
        void GetAt(uint i, out PKey k);
        void GetValue(ref PKey k, out PVal v);
        void SetValue(ref PKey k, ref PVal v);
        void Commit();
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    public class LinkCoClass { }

    public static class Linker {
        public static void Build(string lnkPath, string target, string arguments, string workDir,
                                 string desc, string iconPath, string aumid) {
            object o = new LinkCoClass();
            var link = (IShellLinkW)o;
            link.SetPath(target);
            if (arguments != null) link.SetArguments(arguments);
            if (workDir != null) link.SetWorkingDirectory(workDir);
            if (desc != null) link.SetDescription(desc);
            if (iconPath != null) link.SetIconLocation(iconPath, 0);

            var st = (IPropStore)o;
            var k = new PKey();
            k.fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"); // PKEY_AppUserModel_ID
            k.pid = 5;
            var v = new PVal();
            v.vt = 31; // VT_LPWSTR
            v.p = Marshal.StringToCoTaskMemUni(aumid);
            try {
                st.SetValue(ref k, ref v);
                st.Commit();
            } finally {
                Marshal.FreeCoTaskMem(v.p);
            }
            ((IPersistFile)o).Save(lnkPath, true);
            Marshal.ReleaseComObject(o);
        }
    }
}
'@

function Get-StartMenuLinkPath {
    return (Join-Path (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs') ($script:ToastAppName + '.lnk'))
}

function Save-AppIcon {
    # 给快捷方式和通知用的图标:32x32 PNG 塞进 ICO 容器(Vista 以后都认)
    param([string]$Path)
    try {
        Add-Type -AssemblyName System.Drawing
        $bmp = [System.Drawing.Bitmap]::new(32, 32)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)
        $g.FillEllipse([System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(46, 160, 67)), 2, 2, 28, 28)
        $g.DrawEllipse([System.Drawing.Pen]::new([System.Drawing.Color]::White, 2), 3, 3, 26, 26)
        $g.DrawString('网', [System.Drawing.Font]::new('Microsoft YaHei', 13, [System.Drawing.FontStyle]::Bold), [System.Drawing.Brushes]::White, 3, 5)
        $g.Dispose()

        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $png = $ms.ToArray()
        $ms.Dispose()
        $bmp.Dispose()

        $fs = [System.IO.File]::Create($Path)
        $bw = New-Object System.IO.BinaryWriter($fs)
        $bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]1)
        $bw.Write([byte]32); $bw.Write([byte]32); $bw.Write([byte]0); $bw.Write([byte]0)
        $bw.Write([uint16]1); $bw.Write([uint16]32)
        $bw.Write([uint32]$png.Length); $bw.Write([uint32]22)
        $bw.Write($png)
        $bw.Flush(); $bw.Close()
        return $true
    } catch { return $false }
}

function Register-CampusProtocol {
    # campusnet: 协议 —— 通知上的按钮点下去就执行
    try {
        $lines = @(
            "' Auto-generated by Connect-CampusNet.ps1 - campusnet: protocol handler",
            'Set sh = CreateObject("WScript.Shell")',
            'uri = LCase(WScript.Arguments(0))',
            "' compare the part after the colon exactly (InStr would let log match login)",
            'p = uri',
            'i = InStr(p, ":")',
            'If i > 0 Then p = Mid(p, i + 1)',
            'If Right(p, 1) = "/" Then p = Left(p, Len(p) - 1)',
            'mode = "menu"',
            'window = 1',
            'If p = "login" Then mode = "login" : window = 0',
            'If p = "repair" Then mode = "repair" : window = 0',
            'If p = "status" Then mode = "status" : window = 0',
            'If p = "log" Then mode = "log"',
            ('sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""' + $SelfScript + '"" -Mode " & mode, window, False')
        )
        Set-Content -Path $script:UriVbs -Value $lines -Encoding ASCII

        $root = 'HKCU:\Software\Classes\campusnet'
        New-Item -Path $root -Force | Out-Null
        New-ItemProperty -Path $root -Name '(default)' -Value 'URL:CampusNet' -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $root -Name 'URL Protocol' -Value '' -PropertyType String -Force | Out-Null
        $cmd = 'HKCU:\Software\Classes\campusnet\shell\open\command'
        New-Item -Path $cmd -Force | Out-Null
        $cmdLine = ('wscript.exe //nologo "{0}" "%1"' -f $script:UriVbs)
        New-ItemProperty -Path $cmd -Name '(default)' -Value $cmdLine -PropertyType String -Force | Out-Null
        return $true
    } catch {
        Write-Log ('注册 campusnet: 协议失败:' + $_.Exception.Message) 'WARN'
        return $false
    }
}

function Register-AppIdentity {
    # 让系统认识「校园网自动登录」这个应用:通知中心可用 + 设置里有独立开关
    param([switch]$Force)

    if ($null -ne $script:ToastState -and -not $Force) { return [bool]$script:ToastState }

    try {
        [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    } catch {
        $script:ToastState = $false
        Write-Log ('系统不支持 WinRT 通知,退回托盘气泡:' + $_.Exception.Message) 'WARN'
        return $false
    }

    try {
        if (-not (Test-Path $script:AppIconFile)) { [void](Save-AppIcon -Path $script:AppIconFile) }
        $iconPath = $(if (Test-Path $script:AppIconFile) { $script:AppIconFile } else { 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' })

        $lnk = Get-StartMenuLinkPath
        $need = $true
        if ((Test-Path $lnk) -and -not $Force) {
            try {
                $id = (New-Object -ComObject Shell.Application).Namespace((Split-Path $lnk)).ParseName((Split-Path $lnk -Leaf)).ExtendedProperty('System.AppUserModel.ID')
                if ($id -eq $script:ToastAumid) { $need = $false }
            } catch { }
        }
        if ($need) {
            if (-not ('CampusNetShell.Linker' -as [type])) {
                Add-Type -TypeDefinition $script:LinkHelperCs -Language CSharp
            }
            $target = 'C:\Windows\System32\wscript.exe'
            $args = '"' + $VbsFile + '"'
            if (-not (Test-Path $VbsFile)) {
                # 还没安装过:先用 powershell 直接开托盘,装了之后会自动换成隐藏启动器
                $target = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
                $args = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -Mode tray' -f $SelfScript)
            }
            [CampusNetShell.Linker]::Build($lnk, $target, $args, $WorkDir, $script:ToastAppName, $iconPath, $script:ToastAumid)
            Write-Log ('已注册开始菜单快捷方式(带通知身份):' + $lnk) 'OK'
        }

        # 设置 > 通知 里显示的名字
        $auKey = 'HKCU:\Software\Classes\AppUserModelId\' + $script:ToastAumid
        New-Item -Path $auKey -Force | Out-Null
        New-ItemProperty -Path $auKey -Name 'DisplayName' -Value $script:ToastAppName -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $auKey -Name 'IconUri' -Value $iconPath -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $auKey -Name 'ShowInSettings' -Value 1 -PropertyType DWord -Force | Out-Null
        [void](Register-CampusProtocol)
    } catch {
        Write-Log ('注册通知身份失败(不影响登录):' + $_.Exception.Message) 'WARN'
    }

    $script:ToastState = $true
    return $true
}

function Send-Notify {
    # 统一出口:能用通知中心就发 Toast,不行才退回气泡
    param(
        [string]$Title,
        [string]$Message,
        [ValidateSet('info', 'warn', 'error')]
        [string]$Kind = 'info',
        [string]$Tag = '',
        [string[]]$Actions = @(),
        [int]$MinIntervalSec = 0,
        [switch]$Quiet
    )

    if ($Tag -and $MinIntervalSec -gt 0) {
        $last = $script:ToastSeen[$Tag]
        if ($last -and ((Get-Date) - $last).TotalSeconds -lt $MinIntervalSec) { return $false }
    }

    $sent = $false
    $via = '仅日志'
    if (Register-AppIdentity) {
        try {
            $enc = { param($t) [System.Security.SecurityElement]::Escape([string]$t) }
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.Append('<toast duration="short"><visual><binding template="ToastGeneric">')
            [void]$sb.Append('<text>' + (& $enc $Title) + '</text>')
            [void]$sb.Append('<text>' + (& $enc $Message) + '</text>')
            [void]$sb.Append('</binding></visual>')
            if ($Actions.Count -gt 0) {
                [void]$sb.Append('<actions>')
                foreach ($a in $Actions) {
                    $parts = $a -split '\|', 2
                    $label = $parts[0]
                    $arg = $(if ($parts.Count -gt 1) { $parts[1] } else { 'campusnet:menu' })
                    [void]$sb.Append('<action content="' + (& $enc $label) + '" activationType="protocol" arguments="' + (& $enc $arg) + '"/>')
                }
                [void]$sb.Append('</actions>')
            }
            [void]$sb.Append('</toast>')

            $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
            $xml.LoadXml($sb.ToString())
            $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
            if ($Tag -match '^[A-Za-z0-9]{1,16}$') { $toast.Tag = $Tag }
            [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($script:ToastAumid).Show($toast)
            $sent = $true
            $via = '通知中心'
        } catch {
            Write-Log ('发通知失败(改用气泡):' + $_.Exception.Message) 'WARN'
        }
    }

    if (-not $sent -and $script:TrayNotify) {
        try {
            $ico = switch ($Kind) {
                'warn' { [System.Windows.Forms.ToolTipIcon]::Warning }
                'error' { [System.Windows.Forms.ToolTipIcon]::Error }
                default { [System.Windows.Forms.ToolTipIcon]::Info }
            }
            $script:TrayNotify.ShowBalloonTip(3000, $Title, $Message, $ico)
            $sent = $true
            $via = '气泡'
        } catch { }
    }

    if ($Tag) { $script:ToastSeen[$Tag] = Get-Date }
    if (-not $Quiet) { Write-Log ('通知[' + $via + ']: ' + $Title + ' —— ' + $Message) }

    return $sent
}

# ================= 自动修复(常见毛病的一键处理) =================
# 只做有把握、可逆、不用重启的动作:刷 DNS → 清 Cookie → 重连 WiFi → 校对时间;
# 还不行才动网卡(重置网卡会断网十几秒,所以只有「手动修复」或「连续失败」才做)。
function Repair-FlushDns {
    try {
        $out = & ipconfig /flushdns 2>&1
        if ($LASTEXITCODE -eq 0) { return @{ Ok = $true; Detail = '已清空 DNS 缓存' } }
        $last = @($out) | Select-Object -Last 1
        return @{ Ok = $false; Detail = ('ipconfig 返回 ' + $LASTEXITCODE + ':' + $last) }
    } catch { return @{ Ok = $false; Detail = $_.Exception.Message } }
}

function Repair-ClearCookies {
    $n = 0
    try {
        foreach ($f in @($CookieJar, (Join-Path $WorkDir 'cookies.txt'))) {
            if ($f -and (Test-Path $f)) { Remove-Item $f -Force -ErrorAction Stop; $n++ }
        }
    } catch { }
    if ($n -gt 0) { return @{ Ok = $true; Detail = ('已删除 ' + $n + ' 个 Cookie 文件') } }
    return @{ Ok = $true; Detail = '没有残留 Cookie' }
}

function Repair-Wifi {
    param([string]$Ssid)
    try {
        if (-not $Ssid) { $Ssid = (Get-Settings).wifiSsid }
        if (Test-WifiConnected -Ssid $Ssid) { return @{ Ok = $true; Detail = ('已经连着 ' + $Ssid) } }
        [void](Update-WifiList -Force)
        if (Connect-CampusWifi -Ssid $Ssid) { return @{ Ok = $true; Detail = ('已连上 ' + $Ssid) } }
        return @{ Ok = $false; Detail = ('连不上 ' + $Ssid + '(不在信号范围?)') }
    } catch { return @{ Ok = $false; Detail = $_.Exception.Message } }
}

function Repair-Clock {
    try {
        if (Sync-Clock -Always) { return @{ Ok = $true; Detail = '时间已校对' } }
        return @{ Ok = $false; Detail = '校对不了(需要管理员或门户时间取不到)' }
    } catch { return @{ Ok = $false; Detail = $_.Exception.Message } }
}

function Repair-Nic {
    param([string]$Name)
    try {
        if (-not $Name) {
            $w = Get-WifiInfo
            $Name = $w.Interface
        }
        if (-not $Name) { return @{ Ok = $false; Detail = '找不到网卡名' } }
        $ad = Get-NetAdapter -Name $Name -ErrorAction SilentlyContinue
        if (-not $ad) {
            $ad = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' })[0]
        }
        if (-not $ad) { return @{ Ok = $false; Detail = ('系统里没有网卡 ' + $Name) } }

        if (-not (Test-Elevated)) {
            # 不是管理员:交给提权助手任务做(安装时用管理员注册过才有,平时不跑)
            $helper = $null
            try { $helper = Get-ScheduledTask -TaskName $RepairTaskName -ErrorAction SilentlyContinue } catch { }
            if (-not $helper) {
                return @{ Ok = $false; NeedAdmin = $true; Detail = '重置网卡要管理员:用管理员运行一次安装(加 -Elevated)之后就能自动做' }
            }
            Write-Log '当前不是管理员,改用提权助手任务重置网卡…' 'WARN'
            try {
                Start-ScheduledTask -TaskName $RepairTaskName -ErrorAction Stop
                for ($i = 0; $i -lt 45; $i++) {
                    Start-Sleep -Seconds 1
                    $st = (Get-NetAdapter -Name $ad.Name -ErrorAction SilentlyContinue).Status
                    if ($st -eq 'Up') { return @{ Ok = $true; Detail = ('已由提权任务重置 ' + $ad.Name) } }
                }
                return @{ Ok = $false; NeedAdmin = $true; Detail = '提权任务跑过了,但网卡 45 秒内没回到 Up' }
            } catch {
                return @{ Ok = $false; NeedAdmin = $true; Detail = ('提权任务启动失败:' + $_.Exception.Message) }
            }
        }

        Write-Log ('重置网卡 ' + $ad.Name + '(会短暂断网)…') 'WARN'
        Restart-NetAdapter -Name $ad.Name -Confirm:$false -ErrorAction Stop
        for ($i = 0; $i -lt 30; $i++) {
            Start-Sleep -Seconds 1
            $st = (Get-NetAdapter -Name $ad.Name -ErrorAction SilentlyContinue).Status
            if ($st -eq 'Up') { break }
        }
        $now = (Get-NetAdapter -Name $ad.Name -ErrorAction SilentlyContinue).Status
        if ($now -eq 'Up') { return @{ Ok = $true; Detail = ('已重置 ' + $ad.Name + ',网卡回到 Up') } }
        return @{ Ok = $false; Detail = ('已重置 ' + $ad.Name + ',但状态还是 ' + $now) }
    } catch { return @{ Ok = $false; Detail = $_.Exception.Message } }
}

function Invoke-Repair {
    param(
        [ValidateSet('safe', 'full')]
        [string]$Level = 'safe',
        [switch]$Auto,
        [switch]$Force,
        [string]$Why = '',
        [string]$NicName = ''
    )

    $res = [pscustomobject]@{
        Ran       = $false
        Steps     = New-Object System.Collections.ArrayList
        Online    = $false
        Note      = ''
        NeedAdmin = $false
    }

    $set = Get-Settings
    $now = Get-Date
    if ($Auto) {
        if ($set.lastRepair) {
            try {
                if (($now - [datetime]::ParseExact($set.lastRepair, 'yyyy-MM-dd HH:mm:ss', $null)).TotalSeconds -lt 300) {
                    $res.Note = '刚修过,5 分钟内不重复'
                    return $res
                }
            } catch { }
        }
        if ($Level -eq 'full' -and $set.lastNicReset) {
            try {
                if (($now - [datetime]::ParseExact($set.lastNicReset, 'yyyy-MM-dd HH:mm:ss', $null)).TotalSeconds -lt 1800) {
                    $Level = 'safe'
                    $res.Note = '重置网卡 30 分钟内只做一次,本次先做轻量修复'
                }
            } catch { }
        }
    }

    if (-not $Auto -and -not $Force) {
        Write-Stdout ''
        Write-Stdout ('将依次尝试:刷新 DNS 缓存、清空门户 Cookie' + $(if ($script:AutoWifi -and -not $NoWifi) { '、重新连接 WiFi' } else { '' }) + '、校对系统时间' + $(if ($Level -eq 'full') { '、重置网卡(会断网十几秒)' } else { '' }))
        $ans = Read-Host '开始修复? (Y/n)'
        if ($ans -match '^[nN]') { $res.Note = '已取消'; Write-Stdout '已取消'; return $res }
    }

    $res.Ran = $true
    Write-Log ('开始修复(' + $Level + ')' + $(if ($Why) { ' 原因: ' + $Why } else { '' }) + ' …') 'WARN'

    $plan = New-Object System.Collections.ArrayList
    [void]$plan.Add(@{ Name = '刷新 DNS 缓存'; Do = { Repair-FlushDns } })
    [void]$plan.Add(@{ Name = '清空门户 Cookie'; Do = { Repair-ClearCookies } })
    if ($script:AutoWifi -and -not $NoWifi) { [void]$plan.Add(@{ Name = '重新连接 WiFi'; Do = { Repair-Wifi -Ssid $WifiSsid } }) }
    [void]$plan.Add(@{ Name = '校对系统时间'; Do = { Repair-Clock } })
    if ($Level -eq 'full') { [void]$plan.Add(@{ Name = '重置网卡'; Do = { Repair-Nic -Name $NicName } }) }

    foreach ($step in $plan) {
        $r = $null
        try { $r = & $step.Do } catch { $r = @{ Ok = $false; Detail = $_.Exception.Message } }
        $ok = [bool]$r.Ok
        if ($r.NeedAdmin) { $res.NeedAdmin = $true }
        $detail = [string]$r.Detail
        [void]$res.Steps.Add([pscustomobject]@{ Name = $step.Name; Ok = $ok; Detail = $detail })
        if ($ok) { Write-Log ('修复 · ' + $step.Name + ' → ' + $detail) 'OK' }
        else { Write-Log ('修复 · ' + $step.Name + ' → 没成功:' + $detail) 'WARN' }

        Start-Sleep -Milliseconds 300
        if (Test-Online) { break }
    }

    $res.Online = [bool]((Get-NetStatus).Online)

    $set = Get-Settings
    $set.lastRepair = $now.ToString('yyyy-MM-dd HH:mm:ss')
    if (@($res.Steps | Where-Object { $_.Name -eq '重置网卡' }).Count -gt 0) { $set.lastNicReset = $set.lastRepair }
    try { Save-Settings -Settings $set } catch { }

    if ($res.Online) { Write-Log ('修复结束:已恢复在线(共 ' + $res.Steps.Count + ' 步)') 'OK' }
    else { Write-Log ('修复结束:仍然不通(共 ' + $res.Steps.Count + ' 步)') 'WARN' }
    return $res
}

# ================= 托盘 =================
function Get-WatcherProcess {
    foreach ($p in (Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue)) {
        if ($p.ProcessId -ne $PID -and $p.CommandLine -match 'Connect-CampusNet\.ps1' -and $p.CommandLine -match '\-Mode\s+watch') {
            return $p
        }
    }
    return $null
}

function Start-WatcherIfNeeded {
    if (Get-WatcherProcess) { return }
    Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden `
        -ArgumentList ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Mode watch -Background' -f $SelfScript)
    Write-Log '已拉起后台守护进程'
}
function Start-TrayIfNeeded {
    # 菜单关掉之后把托盘拉回来:不然用户会觉得"窗口一关图标就没了"
    if (Get-TrayProcess) { return $false }
    try {
        if (Test-Path $VbsFile) {
            Start-Process -FilePath 'wscript.exe' -ArgumentList ('//nologo "{0}"' -f $VbsFile) -WindowStyle Hidden
        } else {
            Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden `
                -ArgumentList ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Mode tray -Background' -f $SelfScript)
        }
        Write-Log '已拉起托盘进程'
        return $true
    } catch {
        Write-Log ('拉起托盘失败:' + $_.Exception.Message) 'WARN'
        return $false
    }
}

function Get-TrayProcess {
    foreach ($p in (Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue)) {
        if ($p.ProcessId -ne $PID -and $p.CommandLine -match 'Connect-CampusNet\.ps1' -and $p.CommandLine -match '\-Mode\s+tray') { return $p }
    }
    return $null
}

function Stop-Watcher {
    $p = Get-WatcherProcess
    if ($p) {
        try {
            Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
            Write-Log ('已停止守护进程 pid=' + $p.ProcessId) 'OK'
        } catch { Write-Log ('停止守护失败:' + $_.Exception.Message) 'WARN' }
    }
}

function Stop-Tray {
    $found = $false
    foreach ($p in (Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue)) {
        if ($p.ProcessId -ne $PID -and $p.CommandLine -match 'Connect-CampusNet\.ps1' -and $p.CommandLine -match '\-Mode\s+tray') {
            try {
                Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
                Write-Log ('已停止托盘 pid=' + $p.ProcessId) 'OK'
                $found = $true
            } catch { }
        }
    }
    return $found
}

function New-StatusIcon {
    param([string]$Kind)

    $c = switch ($Kind) {
        'green'  { [System.Drawing.Color]::FromArgb(46, 160, 67) }
        'red'    { [System.Drawing.Color]::FromArgb(200, 50, 50) }
        'orange' { [System.Drawing.Color]::FromArgb(230, 140, 40) }
        default  { [System.Drawing.Color]::FromArgb(120, 120, 120) }
    }

    $bmp = [System.Drawing.Bitmap]::new(32, 32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.FillEllipse([System.Drawing.SolidBrush]::new($c), 2, 2, 28, 28)
    $g.DrawEllipse([System.Drawing.Pen]::new([System.Drawing.Color]::White, 2), 3, 3, 26, 26)
    $g.DrawString('网', [System.Drawing.Font]::new('Microsoft YaHei', 13, [System.Drawing.FontStyle]::Bold), [System.Drawing.Brushes]::White, 3, 5)
    $g.Dispose()
    return [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

function Update-TrayState {
    if (-not $script:TrayNotify) { return }

    # 网速常驻采样:内部自带基线,同一秒内重复调用会复用上次结果
    $sp = Get-NetSpeed

    $st = Read-State
    # Get-CimInstance 查守护进程要几十毫秒,每 2 秒查一次太浪费 —— 缓存 10 秒
    if ($script:TrayAliveAt -eq [datetime]::MinValue -or ((Get-Date) - $script:TrayAliveAt).TotalSeconds -ge 10) {
        $script:TrayAlive = [bool](Get-WatcherProcess)
        $script:TrayAliveAt = Get-Date
    }
    $alive = [bool]$script:TrayAlive

    $kind = 'gray'
    $tip = '校园网:守护未运行'
    $statusText = '状态:守护未运行'
    $accountText = '账号:—'

    if ($st) {
        $age = Get-StateAge -State $st
        $fresh = $age -lt ($IntervalSec * 3)
        $accountText = '账号:' + $(if ($st.account) { $st.account } else { '—' })

        if ($st.note -eq 'logging-in' -and $fresh) {
            $kind = 'orange'; $tip = '校园网:正在登录…'; $statusText = ('状态:正在登录  (' + $st.timestamp + ')')
        } elseif ($st.online -and $fresh) {
            $kind = 'green'; $tip = '校园网:在线'; $statusText = ('状态:在线  (' + $st.timestamp + ')')
        } elseif ($st.note -eq 'no-credential') {
            $kind = 'gray'; $tip = '校园网:未设置账号'; $statusText = '状态:未设置账号'
        } elseif ($fresh) {
            $kind = 'red'; $tip = '校园网:离线'; $statusText = ('状态:离线  (' + $st.timestamp + ')')
        } else {
            $kind = 'gray'; $tip = '校园网:状态过期'; $statusText = ('状态:状态过期  (' + $st.timestamp + ')')
        }
        if (-not $alive -and $st.online) { $tip += '(守护已退出)' }
    }

    if ($kind -ne $script:TrayKind) {
        $script:TrayKind = $kind
        $script:TrayNotify.Icon = $script:Icons[$kind]
    }
    # 鼠标悬停的提示里带上实时网速(只在线时显示),太长就截断
    if ($sp -and $sp.Ready -and $kind -eq 'green') { $tip += '  ↓' + $sp.DownShort + ' ↑' + $sp.UpShort }
    if ($tip.Length -gt 63) { $tip = $tip.Substring(0, 63) }
    if ($script:TrayNotify.Text -ne $tip) { $script:TrayNotify.Text = $tip }
    if ($script:TrayItemStatus) { $script:TrayItemStatus.Text = $statusText }
    if ($script:TrayItemAccount) { $script:TrayItemAccount.Text = $accountText }

    if ($script:TrayItemSpeed) {
        $spText = '网速: ↓ ' + $sp.DownText + '   ↑ ' + $sp.UpText
        if ($script:TrayItemSpeed.Text -ne $spText) { $script:TrayItemSpeed.Text = $spText }
    }
    if ($script:TrayItemTotal) {
        $totText = '峰值 ↓ ' + $sp.PeakDownText + '  ↑ ' + $sp.PeakUpText + '   累计 ' + $sp.TotalText
        if ($script:TrayItemTotal.Text -ne $totText) { $script:TrayItemTotal.Text = $totText }
    }

    # 状态变化时发通知(第一次读取不发,免得开机就弹)
    $nowOnline = ($kind -eq 'green')
    if ($null -ne $script:LastOnline -and $nowOnline -ne $script:LastOnline) {
        if ($nowOnline) {
            [void](Send-Notify -Title '校园网已连接' -Message '认证成功,可以上网了' `
                    -Tag 'online' -MinIntervalSec 20 -Actions @('打开菜单|campusnet:menu'))
        } else {
            [void](Send-Notify -Title '校园网已断开' -Message '正在自动重连…' -Kind 'warn' `
                    -Tag 'offline' -MinIntervalSec 20 -Actions @('立即登录|campusnet:login', '查看日志|campusnet:log'))
        }
    }
    $script:LastOnline = $nowOnline
}

function Invoke-Tray {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # 托盘里的 UI 回调出异常时只记日志,不要直接退出(否则用户看到的是"图标自己没了")
    try {
        [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
        [System.Windows.Forms.Application]::add_ThreadException({
                param($sender, $e)
                try { Write-Log ('托盘异常(已忽略,继续运行):' + $e.Exception.Message) 'WARN' } catch { }
            })
    } catch { }

    # 托盘一起来就把通知身份备好:之后的通知都进通知中心(而不是被吞掉的气泡)
    [void](Register-AppIdentity)

    $trayMutex = New-Object System.Threading.Mutex($false, 'Local\CampusNetTray')
    if (-not $trayMutex.WaitOne(0)) {
        Write-Log '托盘已经在运行了' 'WARN'
        return
    }

    # 图标只生成一次,之后只切换,避免长期运行泄漏 GDI 句柄
    $script:Icons = @{
        green  = New-StatusIcon 'green'
        red    = New-StatusIcon 'red'
        orange = New-StatusIcon 'orange'
        gray   = New-StatusIcon 'gray'
    }
    $script:TrayKind = ''

    $script:TrayNotify = New-Object System.Windows.Forms.NotifyIcon
    $script:TrayNotify.Icon = $script:Icons['gray']
    $script:TrayNotify.Text = '校园网自动登录'
    $script:TrayNotify.Visible = $true

    $menu = New-Object System.Windows.Forms.ContextMenuStrip

    # WiFi 开关放最上面:优先级最高
    $itemAutoWifi = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemAutoWifi.Text = 'WiFi 自动连接: ' + $(if ($script:AutoWifi) { '开' } else { '关' }) + '  (' + $WifiSsid + ')'
    $itemAutoWifi.add_Click({
            $s2 = Get-Settings
            $s2.autoWifi = -not [bool]$s2.autoWifi
            Save-Settings -Settings $s2
            $script:AutoWifi = [bool]$s2.autoWifi
            $itemAutoWifi.Text = 'WiFi 自动连接: ' + $(if ($script:AutoWifi) { '开' } else { '关' }) + '  (' + $s2.wifiSsid + ')'
            Write-Log ('托盘:WiFi 自动连接已' + $(if ($script:AutoWifi) { '开启' } else { '关闭' })) 'OK'
            if ($script:AutoWifi) { [void](Connect-CampusWifi -Ssid $s2.wifiSsid) }
        })
    [void]$menu.Items.Add($itemAutoWifi)
    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $script:TrayItemStatus = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:TrayItemStatus.Text = '状态:读取中…'
    $script:TrayItemStatus.Enabled = $false
    $script:TrayItemAccount = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:TrayItemAccount.Text = '账号:—'
    $script:TrayItemAccount.Enabled = $false
    [void]$menu.Items.Add($script:TrayItemStatus)
    [void]$menu.Items.Add($script:TrayItemAccount)

    # 网速常驻显示,没有开关
    $script:TrayItemSpeed = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:TrayItemSpeed.Text = '网速: 采样中…'
    $script:TrayItemSpeed.Enabled = $false
    $script:TrayItemTotal = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:TrayItemTotal.Text = '峰值 / 累计'
    $script:TrayItemTotal.Enabled = $false
    [void]$menu.Items.Add($script:TrayItemSpeed)
    [void]$menu.Items.Add($script:TrayItemTotal)
    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $itemLogin = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemLogin.Text = '立即登录'
    $itemLogin.add_Click({
            $cred = Get-ActiveCred
            if (-not $cred) { [void][System.Windows.Forms.MessageBox]::Show('还没有保存账号,请用「打开设置菜单」添加', '校园网'); return }
            [void](Invoke-LoginOnce -Cred $cred)
            Update-TrayState
        })
    [void]$menu.Items.Add($itemLogin)

    $itemSwitch = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemSwitch.Text = '切换到下一个账号'
    $itemSwitch.add_Click({
            $store = Get-Store
            $list = @($store.accounts)
            if ($list.Count -lt 2) { [void][System.Windows.Forms.MessageBox]::Show('至少要有两个已保存的账号才能切换', '校园网'); return }
            $curIdx = 0
            for ($i = 0; $i -lt $list.Count; $i++) { if ($list[$i].user -eq $store.active) { $curIdx = $i } }
            $next = $list[($curIdx + 1) % $list.Count]
            if (Set-ActiveAccount -Account $next.user) {
                Stop-Watcher
                Start-Sleep -Milliseconds 500
                Start-WatcherIfNeeded
                [void][System.Windows.Forms.MessageBox]::Show(('已切换到 ' + $next.user + ',守护已用新账号重启'), '校园网')
            }
            Update-TrayState
        })
    [void]$menu.Items.Add($itemSwitch)

    $itemSync = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemSync.Text = '校正系统时间'
    $itemSync.add_Click({ [void](Sync-Clock -Always) })
    [void]$menu.Items.Add($itemSync)

    $itemWifi = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemWifi.Text = ('连接 WiFi(' + $WifiSsid + ')')
    $itemWifi.add_Click({ [void](Connect-CampusWifi -Ssid $WifiSsid) })
    [void]$menu.Items.Add($itemWifi)

    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $itemLog = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemLog.Text = '查看日志'
    $itemLog.add_Click({ if (Test-Path $LogFile) { Start-Process 'notepad.exe' ('"{0}"' -f $LogFile) } })
    [void]$menu.Items.Add($itemLog)

    $itemDoctor = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemDoctor.Text = '一键诊断(生成报告)'
    $itemDoctor.add_Click({
            [void](Invoke-Doctor)
            $rp = Join-Path $WorkDir 'doctor-report.txt'
            if (Test-Path $rp) { Start-Process 'notepad.exe' ('"{0}"' -f $rp) }
        })
    [void]$menu.Items.Add($itemDoctor)

    $itemRepair = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemRepair.Text = '一键修复(DNS / Cookie / 网卡)'
    $itemRepair.add_Click({
            Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Mode repair -Force -RepairLevel full' -f $SelfScript)
            [void](Send-Notify -Title '开始一键修复' -Message '正在刷新 DNS、清空 Cookie、必要时重置网卡…' -Tag 'repairstart' -MinIntervalSec 30)
        })
    [void]$menu.Items.Add($itemRepair)

    $itemNotifyTest = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemNotifyTest.Text = '发送测试通知'
    $itemNotifyTest.add_Click({
            [void](Send-Notify -Title '校园网自动登录' -Message '这是一条测试通知,点下面的按钮可以打开菜单' `
                    -Actions @('打开菜单|campusnet:menu', '立即登录|campusnet:login'))
        })
    [void]$menu.Items.Add($itemNotifyTest)
    # ---- 内存(Mem Reduct 精简版:轻清 / 系统清理 / 自动阈值清理)----
    $itemMem = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMem.Text = '内存'

    $itemMemState = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemState.Text = '可用 - / - MB'
    $itemMemState.add_Click({
            try {
                $m = Get-MemorySnapshot
                $itemMemState.Text = ('可用 {0} MB / {1} MB ({2}%)' -f $m.FreeMB, $m.TotalMB, $m.FreePct)
            } catch { }
        })
    [void]$itemMem.DropDownItems.Add($itemMemState)

    $itemMemLight = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemLight.Text = '轻清(自家进程,免管理员)'
    $itemMemLight.add_Click({
            try {
                $r = Invoke-MemoryClean -Scope light
                $m = Get-MemorySnapshot
                $itemMemState.Text = ('可用 {0} MB / {1} MB ({2}%)' -f $m.FreeMB, $m.TotalMB, $m.FreePct)
                [void][System.Windows.Forms.MessageBox]::Show(($r.Parts | ForEach-Object { $_.Message }) -join "`r`n", '轻清完成')
            } catch { }
        })
    [void]$itemMem.DropDownItems.Add($itemMemLight)

    $itemMemSys = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemSys.Text = '系统清理(standby + 文件缓存)'
    $itemMemSys.add_Click({
            try {
                $r = Invoke-MemoryClean -Scope system
                $ok = $true
                foreach ($p in $r.Parts) { if (-not $p.Ok) { $ok = $false } }
                if ($ok) {
                    [void][System.Windows.Forms.MessageBox]::Show(($r.Parts | ForEach-Object { $_.Message }) -join "`r`n", '系统清理')
                } elseif (Start-MemCleanTask) {
                    [void](Send-Notify -Title '内存清理' -Message '直接清理被拒,已改由提权任务执行,2 秒后可在日志里看结果' -Kind 'info' -Tag 'memclean' -MinIntervalSec 10)
                } else {
                    [void][System.Windows.Forms.MessageBox]::Show('系统清理被系统拒绝,且没有提权任务可用。' + "`r`n" + '请以管理员身份运行安装(加 -Elevated)注册提权清理任务,或直接以管理员运行本脚本。', '系统清理不可用')
                }
                $m = Get-MemorySnapshot
                $itemMemState.Text = ('可用 {0} MB / {1} MB ({2}%)' -f $m.FreeMB, $m.TotalMB, $m.FreePct)
            } catch { }
        })
    [void]$itemMem.DropDownItems.Add($itemMemSys)

    $itemMemAll = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemAll.Text = '全部清理(轻清 + 系统)'
    $itemMemAll.add_Click({
            try {
                $r = Invoke-MemoryClean -Scope all
                $ok = $true
                foreach ($p in $r.Parts) { if (-not $p.Ok) { $ok = $false } }
                if ($ok) {
                    [void][System.Windows.Forms.MessageBox]::Show(($r.Parts | ForEach-Object { $_.Message }) -join "`r`n", '全部清理')
                } elseif (Start-MemCleanTask) {
                    [void](Send-Notify -Title '内存清理' -Message '直清部分被拒,已改由提权任务执行(含系统缓存)' -Kind 'info' -Tag 'memclean' -MinIntervalSec 10)
                } else {
                    [void][System.Windows.Forms.MessageBox]::Show('清理被系统拒绝,且没有提权任务可用。可以先点「轻清」,或管理员运行安装注册提权清理任务。', '全部清理')
                }
                $m = Get-MemorySnapshot
                $itemMemState.Text = ('可用 {0} MB / {1} MB ({2}%)' -f $m.FreeMB, $m.TotalMB, $m.FreePct)
            } catch { }
        })
    [void]$itemMem.DropDownItems.Add($itemMemAll)

    [void]$itemMem.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $itemMemAuto = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemAuto.Text = '内存紧张时自动清理'
    $itemMemAuto.Checked = [bool](Get-Settings).memCleanAuto
    $itemMemAuto.add_Click({
            try {
                $s = Get-Settings
                $s.memCleanAuto = -not [bool]$s.memCleanAuto
                Save-Settings -Settings $s
                $itemMemAuto.Checked = [bool]$s.memCleanAuto
                Write-Log ('内存自动清理:' + $(if ($s.memCleanAuto) { '开' } else { '关' }))
            } catch { }
        })
    [void]$itemMem.DropDownItems.Add($itemMemAuto)

    # 阈值可手动调:可用内存低于该百分比时触发
    $itemMemPct = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemPct.Text = ('触发阈值: 可用 {0}% 以下' -f [int](Get-Settings).memCleanMinFreePct)
    foreach ($pv in @(5, 8, 10, 12, 15, 20, 25)) {
        $miPv = New-Object System.Windows.Forms.ToolStripMenuItem
        $miPv.Text = ('可用 {0}% 以下' -f $pv)
        $miPv.Tag = $pv
        $miPv.Checked = ([int](Get-Settings).memCleanMinFreePct -eq $pv)
        $miPv.add_Click({
                try {
                    $s = Get-Settings
                    $s.memCleanMinFreePct = [int]$this.Tag
                    Save-Settings -Settings $s
                    $itemMemPct.Text = ('触发阈值: 可用 {0}% 以下' -f $s.memCleanMinFreePct)
                    foreach ($x in $itemMemPct.DropDownItems) { if ($x.Tag) { $x.Checked = ([int]$x.Tag -eq [int]$s.memCleanMinFreePct) } }
                    Write-Log ('内存自动清理阈值:可用低于 ' + $s.memCleanMinFreePct + '% 时触发')
                } catch { }
            })
        [void]$itemMemPct.DropDownItems.Add($miPv)
    }
    [void]$itemMem.DropDownItems.Add($itemMemPct)

    # 冷却时间可手动调:两次自动清理之间至少隔多久
    $itemMemCool = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemCool.Text = ('冷却时间: {0} 分钟' -f [int](Get-Settings).memCleanCooldownMin)
    foreach ($cv in @(3, 5, 10, 15, 30, 60)) {
        $miCv = New-Object System.Windows.Forms.ToolStripMenuItem
        $miCv.Text = ('{0} 分钟' -f $cv)
        $miCv.Tag = $cv
        $miCv.Checked = ([int](Get-Settings).memCleanCooldownMin -eq $cv)
        $miCv.add_Click({
                try {
                    $s = Get-Settings
                    $s.memCleanCooldownMin = [int]$this.Tag
                    Save-Settings -Settings $s
                    $itemMemCool.Text = ('冷却时间: {0} 分钟' -f $s.memCleanCooldownMin)
                    foreach ($x in $itemMemCool.DropDownItems) { if ($x.Tag) { $x.Checked = ([int]$x.Tag -eq [int]$s.memCleanCooldownMin) } }
                    Write-Log ('内存自动清理冷却:' + $s.memCleanCooldownMin + ' 分钟')
                } catch { }
            })
        [void]$itemMemCool.DropDownItems.Add($miCv)
    }
    [void]$itemMem.DropDownItems.Add($itemMemCool)

    $itemMemSysAuto = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemSysAuto.Text = '自动清理时也清系统缓存(需提权任务)'
    $itemMemSysAuto.Checked = [bool](Get-Settings).memCleanSysInAuto
    $itemMemSysAuto.add_Click({
            try {
                $s = Get-Settings
                $s.memCleanSysInAuto = -not [bool]$s.memCleanSysInAuto
                Save-Settings -Settings $s
                $itemMemSysAuto.Checked = [bool]$s.memCleanSysInAuto
            } catch { }
        })
    [void]$itemMem.DropDownItems.Add($itemMemSysAuto)

    [void]$itemMem.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $itemMemTask = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMemTask.Text = '注册/修复提权清理任务(需管理员)'
    $itemMemTask.add_Click({
            try {
                if (-not (Test-Elevated)) {
                    [void][System.Windows.Forms.MessageBox]::Show('这一步需要管理员权限:请用管理员运行安装(加 -Elevated)。', '注册清理任务')
                    return
                }
                if (Register-MemCleanTask) {
                    [void][System.Windows.Forms.MessageBox]::Show('已注册提权清理任务 CampusNet-MemClean。', '完成')
                } else {
                    [void][System.Windows.Forms.MessageBox]::Show('注册失败,详情见日志。', '失败')
                }
            } catch { }
        })
    [void]$itemMem.DropDownItems.Add($itemMemTask)

    [void]$menu.Items.Add($itemMem)

    # 每次打开菜单都刷新一下可用内存
    $menu.add_Opening({
            try {
                $m = Get-MemorySnapshot
                $itemMemState.Text = ('可用 {0} MB / {1} MB ({2}%)' -f $m.FreeMB, $m.TotalMB, $m.FreePct)
            } catch { }
        })

    $itemMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMenu.Text = '打开设置菜单…'
    $itemMenu.add_Click({
            Start-Process -FilePath 'powershell.exe' `
                -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $SelfScript)
        })
    [void]$menu.Items.Add($itemMenu)

    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $itemExit = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemExit.Text = '退出托盘(同时停止守护)'
    $itemExit.add_Click({
            Stop-Watcher
            $script:TrayNotify.Visible = $false
            $script:TrayNotify.Dispose()
            [System.Windows.Forms.Application]::Exit()
        })
    [void]$menu.Items.Add($itemExit)

    $script:TrayNotify.ContextMenuStrip = $menu

    # 双击托盘 = 立即登录
    $script:TrayNotify.add_DoubleClick({
            $cred = Get-ActiveCred
            if (-not $cred) { [void][System.Windows.Forms.MessageBox]::Show('还没有保存账号,请用「打开设置菜单」添加', '校园网'); return }
            [void](Send-Notify -Title '校园网' -Message '正在登录…' -Tag 'traylogin' -MinIntervalSec 10)
            [void](Invoke-LoginOnce -Cred $cred)
            Update-TrayState
        })

    $script:LastOnline = $null

    # 休眠 / 唤醒:守护是 30 秒一轮,这里抢在它前面立刻补一次认证
    $script:ResumeHandler = {
        param($snd, $evt)
        if ($evt.Mode -eq [Microsoft.Win32.PowerModes]::Resume) {
            try {
                Write-Log '系统从休眠/睡眠唤醒,立刻补一次认证…' 'WARN'
                Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Mode login -Background -Resume' -f $SelfScript)
            } catch {
                Write-Log ('唤醒后补认证没启动:' + $_.Exception.Message) 'WARN'
            }
        }
    }
    try {
        [Microsoft.Win32.SystemEvents]::add_PowerModeChanged($script:ResumeHandler)
        Write-Log '已挂上休眠/唤醒监听'
    } catch {
        Write-Log ('挂休眠/唤醒监听失败(守护仍会在 30 秒内自动重连):' + $_.Exception.Message) 'WARN'
    }

    Start-WatcherIfNeeded
    Update-TrayState

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 2000
    $script:MemAutoCheckAt = [datetime]::MinValue
    $timer.add_Tick({
            try {
                Update-TrayState
                Start-WatcherIfNeeded
                # 内存自动清理:每 60 秒评估一次(阈值/冷却都可在菜单里调)
                $nowT = Get-Date
                if ($script:MemAutoCheckAt -eq [datetime]::MinValue -or ($nowT - $script:MemAutoCheckAt).TotalSeconds -ge 60) {
                    $script:MemAutoCheckAt = $nowT
                    if (Test-MemCleanDue) {
                        $mBefore = Get-MemorySnapshot
                        [void](Invoke-MemoryClean -Scope light)
                        $s = Get-Settings
                        if ($s.memCleanSysInAuto) { [void](Start-MemCleanTask) }
                        $mAfter = Get-MemorySnapshot
                        [void](Send-Notify -Title '内存清理' -Message ('可用内存偏低(' + $mBefore.FreePct + '%),已自动清理;可用 ' + $mBefore.FreeMB + '→' + $mAfter.FreeMB + ' MB') -Kind 'info' -Tag 'memauto' -MinIntervalSec 600)
                    }
                }
            } catch { }
        })
    $timer.Start()

    Write-Log '托盘已启动(右下角「网」图标,右键有菜单)'
    [System.Windows.Forms.Application]::Run()
}

# ================= 安装 / 卸载 =================
function Register-RepairTask {
    # 提权重置网卡用的助手任务:平时不跑,只有自动修复要动网卡时才被唤起(不会弹 UAC)
    if (-not (Test-Elevated)) {
        Write-Log '当前不是管理员:不注册提权修复任务,自动修复时「重置网卡」会跳过。要用请用管理员运行安装(加 -Elevated)' 'WARN'
        return $false
    }
    try {
        $repairArg = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Mode nicreset -Background' -f $SelfScript
        $repairAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $repairArg
        $repairSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
        $repairPrincipal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest
        if (Get-ScheduledTask -TaskName $RepairTaskName -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $RepairTaskName -Confirm:$false
        }
        Register-ScheduledTask -TaskName $RepairTaskName -Action $repairAction -Settings $repairSettings `
            -Principal $repairPrincipal -Description '校园网自动登录:提权重置网卡(供自动修复调用)' | Out-Null
        Write-Log ('已注册提权修复任务 ' + $RepairTaskName + '(自动修复里「重置网卡」可用)') 'OK'
        return $true
    } catch {
        Write-Log ('注册提权修复任务失败(自动修复会跳过重置网卡):' + $_.Exception.Message) 'WARN'
        return $false
    }
}

function Invoke-Install {
    if ((Get-AccountList).Count -eq 0) {
        Write-Log '还没有保存账号。先设置账号密码再安装' 'ERROR'
        return $false
    }

    $runMode = $(if ($NoTray) { 'watch' } else { 'tray' })

    if ($NoVbs) {
        $arg = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -Mode {1} -Background' -f $SelfScript, $runMode
        $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
        Write-Log '使用 powershell -WindowStyle Hidden(启动瞬间可能有极短黑框)' 'WARN'
    } else {
        $vbsLines = @(
            "' Auto-generated by Connect-CampusNet.ps1 - launch hidden",
            'Set sh = CreateObject("WScript.Shell")',
            ('sh.Run "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ""' + $SelfScript + '"" -Mode ' + $runMode + ' -Background", 0, True')
        )
        Set-Content -Path $VbsFile -Value $vbsLines -Encoding ASCII
        $taskAction = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('//nologo "{0}"' -f $VbsFile)
        Write-Log '已生成隐藏启动器 run-hidden.vbs'
    }

    $trigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -ExecutionTimeLimit ([TimeSpan]::Zero)

    $runLevel = 'Limited'
    if ($Elevated) {
        if (Test-Elevated) { $runLevel = 'Highest' }
        else { Write-Log '想以管理员身份运行任务需要先用管理员启动本脚本;本次按普通权限安装(时间偏差只会提醒、不会自动改)' 'WARN' }
    }
    $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel $runLevel

    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }

    Register-ScheduledTask -TaskName $TaskName -Action $taskAction -Trigger $trigger `
        -Settings $settings -Principal $principal `
        -Description '校园网自动登录(掉线自动重连)' | Out-Null

    $t = Get-ScheduledTask -TaskName $TaskName
    Write-Log ('已安装计划任务 ' + $TaskName + '  运行模式:' + $runMode + '  权限:' + $runLevel + '  状态:' + $t.State) 'OK'

    # 提权重置网卡用的助手任务:平时不跑,只有自动修复要动网卡时才被唤起(不会弹 UAC)
    [void](Register-RepairTask)

    # 提权内存清理任务(托盘「系统清理」用,不弹 UAC)
    [void](Register-MemCleanTask)

    Write-Stdout ''
    if ($runMode -eq 'tray') {
        Write-Stdout '说明:'
        Write-Stdout '  · 登录 Windows 后右下角出现「网」图标,右键可立即登录 / 切账号 / 看日志'
        Write-Stdout '  · 守护进程由托盘拉起;托盘退出时会一起停掉'
        Write-Stdout '  · 不想要图标:先卸载,再用 -Mode install -NoTray 重装'
    } else {
        Write-Stdout '说明:登录 Windows 后自动在后台静默运行,没有窗口也没有图标'
    }
    Write-Stdout ('  · 检测间隔 ' + $IntervalSec + ' 秒,每 30 分钟自动校对一次系统时间')
    Write-Stdout ''
    $ans = Read-Host '现在试跑一下吗? (y/N)'
    if ($ans -match '^[yY]') {
        Start-ScheduledTask -TaskName $TaskName
        Start-Sleep -Seconds 6
        Write-Stdout '--- 日志最后 5 行 ---'
        Get-Content $LogFile -Tail 5 -ErrorAction SilentlyContinue
    }
    return $true
}

function Invoke-Uninstall {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Log ('已卸载计划任务 ' + $TaskName) 'OK'
    } else {
        Write-Log '没有找到已安装的计划任务' 'WARN'
    }
    # 顺序很重要:先收托盘,否则托盘的定时器会把守护又拉起来
    [void](Stop-Tray)
    Start-Sleep -Milliseconds 300
    Stop-Watcher
    if (Test-Path $VbsFile) { Remove-Item $VbsFile -Force }
    if (Test-Path $script:UriVbs) { Remove-Item $script:UriVbs -Force }
    if (Get-ScheduledTask -TaskName $RepairTaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $RepairTaskName -Confirm:$false
        Write-Log ('已卸载提权修复任务 ' + $RepairTaskName) 'OK'
    }
}

function Show-Log {
    if (Test-Path $LogFile) {
        Write-Stdout '--- 最近 20 行日志 ---'
        Get-Content $LogFile -Tail 20
    } else { Write-Stdout '还没有日志' }
}

# ================= 远程适配:一键采集 =================
function Invoke-Collect {
    # 给「人不在现场」用的:把判断门户需要的东西全收起来,
    # 文本自动进剪贴板(直接粘到聊天窗口),门户页和它的 JS 打进 zip 一起发。
    param([switch]$NoZip)

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $dir = Join-Path $WorkDir ('collect-' + $stamp)
    $lines = New-Object System.Collections.Generic.List[string]
    $add = { param($t) $lines.Add([string]$t) }

    & $add '===== 校园网远程适配 · 采集报告 ====='
    & $add ('时间: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    & $add ('机器: ' + $env:COMPUTERNAME + ' / 用户 ' + $env:USERNAME)
    & $add ('系统: ' + [Environment]::OSVersion.VersionString + '   PowerShell ' + $PSVersionTable.PSVersion.ToString())
    & $add ('脚本: ' + $SelfScript + '   v' + $script:AppVersion)

    # ---- 无线 ----
    & $add ''
    & $add '--- 无线 ---'
    $ssidNow = ''
    try {
        $w = Get-WifiInfo
        $ssidNow = [string]$w.Ssid
        & $add ('当前: ' + $(if ($w.Ssid) { $w.Ssid } else { '(未连接)' }) + '   状态 ' + $w.State + '   目标 ' + $WifiSsid)
        & $add ('网卡: ' + $w.Interface)
    } catch { & $add ('无线读取失败: ' + $_.Exception.Message) }
    try {
        $nets = @((Update-WifiList).Visible | ForEach-Object { [string]$_.Ssid } | Where-Object { $_ })
        if ($nets.Count) { & $add ('可见网络: ' + (($nets | Select-Object -First 12) -join ' , ')) }
    } catch { }

    # ---- 地址 ----
    & $add ''
    & $add '--- 地址 ---'
    try {
        $cfg = @(Get-NetIPConfiguration -ErrorAction Stop | Where-Object { $_.IPv4DefaultGateway }) | Select-Object -First 1
        if ($cfg) {
            & $add ('IP: ' + (@($cfg.IPv4Address)[0]).IPAddress + '   网关: ' + $cfg.IPv4DefaultGateway.NextHop)
            $dns = @($cfg.DNSServer | Where-Object { $_.AddressFamily -eq 2 } | ForEach-Object { $_.ServerAddresses } | Where-Object { $_ })
            if ($dns.Count) { & $add ('DNS: ' + ($dns -join ' , ')) }
        } else { & $add '没有默认网关(可能没连上)' }
    } catch { & $add ('地址读取失败: ' + $_.Exception.Message) }
    if ($script:ProxyEnvHit -and $script:ProxyEnvHit.Count -gt 0) {
        & $add ('代理变量: ' + ($script:ProxyEnvHit -join ' , ') + '   (已忽略)')
    } else { & $add '代理变量: 无' }

    # ---- 连通性 / 门户探测 ----
    & $add ''
    & $add '--- 探测 ---'
    $ns = Get-NetStatus
    & $add ('联网判定: ' + $(if ($ns.Online) { '在线' } elseif ($ns.Hijacked) { '被劫持' } else { '不通' }) + ' — ' + $ns.Detail)

    $cap = Get-CaptiveRedirect
    & $add ('HTTP 302 门户: ' + $(if ($cap) { $cap } else { '(没抓到)' }))

    $gw = Get-DefaultGatewayIp
    if ($gw) {
        foreach ($gu in @(('http://' + $gw + '/'), ('http://' + $gw + ':8080/'))) {
            $gc = ''
            try { $gc = ((& $Curl -s -o NUL -A $UA -w '%{http_code}' --max-time 4 $gu 2>$null) -join '').Trim() } catch { }
            & $add ('网关 ' + $gu + ' → HTTP ' + $gc)
        }
    }

    $hh = Get-HttpsPortalHint
    if ($hh) {
        & $add ('HTTPS 劫持页: ' + $hh.Title + '   (' + $hh.Size + ' 字节, HTTP ' + $hh.Code + ')')
        if ($hh.Hosts.Count -gt 0) { & $add ('  页面里的域名: ' + ($hh.Hosts -join ' , ')) }
    } else {
        & $add 'HTTPS 劫持页: 没抓到(可能没被拦,或者已经在线)'
    }

    $fp = Get-PortalFingerprint $cap
    & $add ('识别: ' + $fp.Vendor + $(if ($fp.Id) { '(' + $fp.Id + ')' } else { '' }))
    $carrier = Get-CarrierHint -Url $cap -Ssid $ssidNow
    & $add ('运营商线索: ' + $(if ($carrier) { $carrier } else { '无' }))

    # ---- 门户页 + 同站脚本 ----
    & $add ''
    & $add '--- 门户页采集 ---'
    $targets = @()
    if ($cap) { $targets += $cap }
    if ($gw) { $targets += ('http://' + $gw + '/') }
    if ($hh) { foreach ($hv in $hh.Hosts) { $targets += ('http://' + $hv + '/') } }
    # 兜底:当前配置里的门户也试一下,顺手证明它通不通
    $targets += ('http://' + $PortalHost + '/')
    if ($targets.Count -eq 0) { $targets = @('http://' + $PortalHost + '/') }

    try { if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null } } catch { }
    $saved = 0
    foreach ($t in ($targets | Select-Object -Unique)) {
        if ($saved -ge 2) { break }
        if ($t -notmatch '^https?://') { continue }
        $base = ''
        try { $u = [uri]$t; $base = $u.Scheme + '://' + $u.Authority } catch { continue }

        $outFn = Join-Path $dir ('portal-' + ($saved + 1) + '.html')
        $code = ''
        $eff = $t
        try {
            $meta = ((& $Curl -sk -L -A $UA --max-time 10 -o $outFn -w '%{http_code}|%{url_effective}' $t 2>$null) -join '').Trim()
            $mp = $meta -split '\|', 2
            $code = $mp[0]
            if ($mp.Count -gt 1 -and $mp[1] -match '^https?://') { $eff = [string]$mp[1] }
        } catch { }
        if (-not $code -or $code -eq '000' -or -not (Test-Path $outFn)) {
            & $add ($t + ' → HTTP ' + $code + '  抓不到')
            continue
        }
        $html = ''
        try { $html = [System.IO.File]::ReadAllText($outFn, [System.Text.Encoding]::UTF8) } catch { }
        if (-not $html) { & $add ($t + ' → HTTP ' + $code + '  空页面'); continue }

        # 相对路径要按「最终地址」解析,不然 ../share/x.js 会拼错
        $baseUrl = $eff
        try { $bu = [uri]$eff; $base = $bu.Scheme + '://' + $bu.Authority } catch { }

        $title = ''
        if ($html -match '<title>([^<]{1,80})</title>') { $title = $Matches[1].Trim() }
        & $add ($t + ' → HTTP ' + $code + ' , ' + $html.Length + ' 字节 , ' + $(if ($title) { '标题「' + $title + '」' } else { '没有标题' }))
        $saved++

        # 门户的登录逻辑基本都在同站 JS 里,一起取回来
        $jsList = @()
        foreach ($m in [regex]::Matches($html, '(?:src|href)\s*=\s*["'']([^"'']+\.js[^"'']*)["'']')) {
            $src = $m.Groups[1].Value
            if ($src -match '^(data:|javascript:)') { continue }
            $full = $src
            if ($src -notmatch '^https?://') {
                # 用「最终地址」做基准解析,才不会把 ../share/xxx.js 拼错路径
                try { $full = ([Uri]::new([Uri]$baseUrl, $src)).AbsoluteUri }
                catch { $full = $base.TrimEnd('/') + '/' + $src.TrimStart('/') }
            }
            if ($full -like ($base + '*')) { if ($jsList -notcontains $full) { $jsList += $full } }
        }
        $ji = 0
        foreach ($js in ($jsList | Select-Object -First 8)) {
            $ji++
            $fn = Join-Path $dir ('script-' + $saved + '-' + $ji + '.js')
            try { & $Curl -sk -s -A $UA --max-time 8 -o $fn $js 2>$null | Out-Null } catch { }
            if (Test-Path $fn) {
                $sz = (Get-Item $fn).Length
                if ($sz -gt 0) { & $add ('  脚本: ' + (Split-Path $js -Leaf) + '  ' + $sz + ' 字节') }
                else { Remove-Item $fn -Force -ErrorAction SilentlyContinue }
            }
        }
        if ($jsList.Count -gt 8) { & $add ('  (脚本太多,只取了前 8 个)') }
    }
    if ($saved -eq 0) { & $add '一个门户页都没抓到(门户可能只在内网特定条件下才开放)' }

    # ---- 打包 ----
    $zip = ''
    if (-not $NoZip) {
        try {
            $files = @(Get-ChildItem $dir -File -ErrorAction SilentlyContinue)
            if ($files.Count -gt 0) {
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
                $zip = Join-Path $WorkDir ('collect-' + $stamp + '.zip')
                if (Test-Path $zip) { Remove-Item $zip -Force }
                [System.IO.Compression.ZipFile]::CreateFromDirectory($dir, $zip, [System.IO.Compression.CompressionLevel]::Optimal, $false)
            }
        } catch { $zip = ''; & $add ('打包失败: ' + $_.Exception.Message) }
    }

    # ---- 收尾 ----
    & $add ''
    & $add '--- 怎么发回来 ---'
    & $add '1) 上面这些文字已经自动复制到剪贴板,直接粘贴到聊天窗口发出去'
    & $add '2) 同目录下的 collect-<时间>.zip 一起发(里面是门户页和它的 JS,用来还原登录接口)'
    & $add '3) 浏览器弹认证页时,把地址栏那一行也随手截个图'
    if ($zip) { & $add ('压缩包: ' + $zip) }

    $text = ($lines -join "`r`n")
    $txtPath = Join-Path $WorkDir ('collect-' + $stamp + '.txt')
    try { Set-Content -Path $txtPath -Value $text -Encoding UTF8 } catch { }
    try { Set-Clipboard -Value $text } catch { }

    Write-Stdout ''
    Write-Stdout $text
    Write-Stdout ''
    Write-Stdout ('报告文本: ' + $txtPath)
    if ($zip) { Write-Stdout ('压缩包:   ' + $zip) }
    Write-Stdout ''
    Write-Stdout '↑ 文本已经复制到剪贴板,直接粘贴发出去就行'
}

# ================= 新手向导:问一句 + 认门户 =================
function Get-PortalFormDraft {
    # 从门户页的 <form> 里抠出「登录接口 + 字段名」,拼一份能直接用的模板
    param([string]$Html, [string]$PageUrl)

    $d = [pscustomobject]@{
        Ok = $false; LoginUrl = ''; Method = 'POST'
        Template = ''; Encrypt = 'none'; UserField = ''; PwdField = ''; Note = ''
    }
    if ([string]::IsNullOrWhiteSpace($Html)) { return $d }

    $formMatch = $null
    foreach ($fm in [regex]::Matches($Html, '(?is)<form\b[^>]*>(.*?)</form>')) {
        if ($fm.Groups[1].Value -match '(?is)type\s*=\s*["'']?password') { $formMatch = $fm; break }
    }
    if (-not $formMatch) { return $d }

    $open = $formMatch.Value
    $inner = $formMatch.Groups[1].Value

    $action = ''
    $m = [regex]::Match($open, '(?is)\baction\s*=\s*["'']([^"'']*)["'']')
    if ($m.Success) { $action = $m.Groups[1].Value }

    $method = 'POST'
    $m2 = [regex]::Match($open, '(?is)\bmethod\s*=\s*["'']([^"'']*)["'']')
    if ($m2.Success -and $m2.Groups[1].Value -match '(?i)get') { $method = 'GET' }

    $loginUrl = $PageUrl
    if ($action -and $action -notmatch '^(?i)javascript:|^#') {
        if ($action -match '^(?i)https?://') { $loginUrl = $action }
        else { try { $loginUrl = ([Uri]::new([Uri]$PageUrl, $action)).AbsoluteUri } catch { } }
    }

    $parts = @()
    foreach ($im in [regex]::Matches($inner, '(?is)<input\b[^>]*>')) {
        $tag = $im.Value
        $nm = [regex]::Match($tag, '(?is)\bname\s*=\s*["'']([^"'']+)["'']')
        if (-not $nm.Success) { continue }
        $name = $nm.Groups[1].Value

        $ty = 'text'
        $t2 = [regex]::Match($tag, '(?is)\btype\s*=\s*["'']?([a-zA-Z]+)')
        if ($t2.Success) { $ty = $t2.Groups[1].Value.ToLowerInvariant() }

        $val = ''
        $v2 = [regex]::Match($tag, '(?is)\bvalue\s*=\s*["'']([^"'']*)["'']')
        if ($v2.Success) { $val = $v2.Groups[1].Value }

        if ($ty -eq 'password') {
            $parts += ($name + '={pwd}')
            $d.PwdField = $name
        } elseif ((-not $d.UserField) -and ($ty -match '^(text|tel|number|email|search)$') -and
                  ($name -match '(?i)user|account|number|num|mobile|phone|card|login|name|id')) {
            $parts += ($name + '={user}')
            $d.UserField = $name
        } elseif ($name -and $val) {
            $parts += ($name + '=' + [uri]::EscapeDataString($val))
        }
    }
    if (-not $d.UserField -or -not $d.PwdField -or $parts.Count -lt 2) { return $d }

    $d.LoginUrl = $loginUrl
    $d.Method = $method
    $d.Template = ($parts -join '&')
    $d.Ok = $true

    # 猜加密(只是猜:看页面/内联脚本里有没有 md5( / rc4 / base64)
    if ($Html -match '(?i)\bmd5\s*\(') { $d.Encrypt = 'md5' }
    elseif ($Html -match '(?i)\brc4\b') { $d.Encrypt = 'rc4_ms' }
    elseif ($Html -match '(?i)\bbase64\b') { $d.Encrypt = 'base64' }

    $d.Note = ('从门户表单自动生成:用户字段 ' + $d.UserField + ',密码字段 ' + $d.PwdField + ',加密按页面特征猜为 ' + $d.Encrypt)
    return $d
}

function Invoke-Wizard {
    # 第一次用:问一句「这个 WiFi 是不是校园网」,不是就换名字;然后认门户、生成配置
    Write-Stdout ''
    Write-Stdout '=================================================='
    Write-Stdout '   新手向导:认门户 + 设 WiFi'
    Write-Stdout '=================================================='

    # ---- 1. WiFi ----
    $wi = $null
    try { $wi = Get-WifiInfo } catch { }
    $cur = $(if ($wi) { [string]$wi.Ssid } else { '' })
    if ($cur) { Write-Stdout ('  当前连的 WiFi: ' + $cur) }
    else { Write-Stdout '  当前没有连着 WiFi(或者用的是有线)' }

    $target = $cur
    if ($cur) {
        $a1 = (Read-Host '  这个网络是「校园网 / 需要自动认证」的吗? [Y/n]').Trim().ToLowerInvariant()
        if ($a1 -eq 'n') { $target = (Read-Host '  要自动连接的校园网 WiFi 名称').Trim() }
    } else {
        $target = (Read-Host '  要自动连接的校园网 WiFi 名称').Trim()
    }

    if ($target) {
        try {
            $s = Get-Settings
            $s.wifiSsid = $target
            $s.autoWifi = $true
            Save-Settings -Settings $s
            $script:WifiSsid = $target
            $script:AutoWifi = $true
            Write-Log ('向导:目标 WiFi = ' + $target + '(自动连接开)') 'OK'
            Write-Stdout ('  ✓ 目标 WiFi 设为: ' + $target)
        } catch {
            Write-Stdout ('  ! 保存 WiFi 设置失败: ' + $_.Exception.Message)
        }
    } else {
        Write-Stdout '  - 没填 WiFi 名称,先用原来的设置'
    }

    # ---- 2. 账号 ----
    $store = Get-Store
    if ($store.accounts.Count -eq 0) {
        $a2 = (Read-Host '  还没有保存账号。现在填吗? [Y/n]').Trim().ToLowerInvariant()
        if ($a2 -ne 'n') {
            $u = (Read-Host '  账号(学号 / 手机号)').Trim()
            if ($u) {
                $pw = Read-PlainPassword '  密码'
                if ($pw) {
                    Set-Account -Account $u -PlainPassword $pw -MakeActive
                    Write-Stdout '  ✓ 账号已保存(DPAPI 加密)'
                }
            }
        }
    } else {
        Write-Stdout ('  已有账号: ' + $store.active + '(' + $store.accounts.Count + ' 个)')
    }

    # ---- 3. 认门户 ----
    Write-Stdout ''
    Write-Stdout '  正在识别门户(最多 15 秒)…'
    $cap = Get-CaptiveRedirect
    $hh = $null
    $gw = Get-DefaultGatewayIp

    if ($cap) {
        Write-Stdout ('  抓到门户跳转: ' + $cap)
    } else {
        if (Test-Online) {
            Write-Stdout '  已经在线,门户不会被劫持(想重认:先注销认证或断网重连再跑一次)'
        } else {
            Write-Stdout '  没抓到 HTTP 跳转,再试试 HTTPS 劫持页和网关…'
        }
        $hh = Get-HttpsPortalHint
        if ($hh) { Write-Stdout ('  HTTPS 劫持页: ' + $hh.Title) }
    }

    $fp = Get-PortalFingerprint $cap
    $applied = $false
    if ($fp.Id) {
        $preset = Get-PortalPreset $fp.Id
        if ($preset) {
            [void](Save-PortalProfile $preset)
            $pf2 = Get-PortalProfile
            if ($fp.Host) { $pf2.host = [string]$fp.Host }
            if ($fp.Id -eq 'esurfing') {
                $pf2.host = 'enet.10000.gd.cn:10001'
                $pf2.page = ''
            } elseif ($fp.Page -and $fp.Page -ne '/') {
                $pf2.page = [string]$fp.Page
            }
            [void](Save-PortalProfile $pf2)
            Write-Stdout ('  ✓ 识别为 ' + $fp.Vendor + ',已套用预设 ' + $fp.Id)
            if (@('srun', 'srun2', 'ruijie') -contains $fp.Id) {
                Write-Stdout '    注意:这个预设是骨架,密码算法可能还要校准(失败就把 collect 包发出来)'
            }
            $applied = $true
        }
    }

    # ---- 4. 没匹配上就从页面表单生成 ----
    if (-not $applied) {
        $pageUrl = ''
        if ($cap) { $pageUrl = $cap }
        elseif ($hh -and @($hh.Hosts).Count -gt 0) { $pageUrl = 'http://' + (@($hh.Hosts)[0]) + '/' }
        elseif ($gw) { $pageUrl = 'http://' + $gw + '/' }
        else { $pageUrl = 'http://' + $PortalHost + '/' }

        Write-Stdout ('  没匹配到内置预设,试着从页面生成配置: ' + $pageUrl)
        $id = [guid]::NewGuid().ToString('N')
        $htmlFn = Join-Path $env:TEMP ('campusnet-wiz-' + $id + '.html')
        $eff = $pageUrl
        $code = ''
        try {
            $meta = ((& $Curl -sk -L -A $UA --max-time 12 -o $htmlFn -w '%{http_code}|%{url_effective}' $pageUrl 2>$null) -join '').Trim()
            $mp = $meta -split '\|', 2
            $code = $mp[0]
            if ($mp.Count -gt 1 -and $mp[1] -match '^https?://') { $eff = [string]$mp[1] }
        } catch { }
        $html = ''
        if ($code -and $code -ne '000' -and (Test-Path -LiteralPath $htmlFn)) {
            try { $html = [System.IO.File]::ReadAllText($htmlFn, [System.Text.Encoding]::UTF8) } catch { }
        }
        Remove-Item -LiteralPath $htmlFn -Force -ErrorAction SilentlyContinue

        if ($html) {
            $draft = Get-PortalFormDraft -Html $html -PageUrl $eff
            if ($draft.Ok) {
                $host2 = ''
                $page2 = '/'
                try { $uu = [uri]$eff; $host2 = [string]$uu.Authority; $page2 = [string]$uu.PathAndQuery } catch { }
                $prof = [pscustomobject]@{
                    id           = 'custom'
                    name         = '自动生成(取自门户表单)'
                    protocol     = 'template'
                    host         = $host2
                    page         = $page2
                    loginUrl     = $draft.LoginUrl
                    method       = $draft.Method
                    contentType  = 'application/x-www-form-urlencoded'
                    encrypt      = $draft.Encrypt
                    template     = $draft.Template
                    successMatch = ''
                    msgMatch     = ''
                    note         = $draft.Note
                }
                if (Save-PortalProfile $prof) {
                    Write-Stdout '  ✓ 已生成门户配置(id=custom):'
                    Write-Stdout ('      ' + $prof.method + ' ' + $prof.loginUrl)
                    Write-Stdout ('      ' + $prof.template)
                    Write-Stdout ('      加密猜测: ' + $prof.encrypt + '(成功判定留空,以连通性为准)')
                    Write-Stdout '    可以先试登录;不行就跑 -Mode portal -DetectPortal 或 -Mode collect'
                    $applied = $true
                }
            } else {
                Write-Stdout '  页面里没找到「密码框 + 账号框」的表单(可能要先进一个页面才出现)'
            }
        } else {
            Write-Stdout ('  页面也打不开(HTTP ' + $code + ')')
        }
    }

    if (-not $applied) {
        Write-Stdout '  - 门户配置没动。可以:①浏览器开一次认证页,把地址栏发出来;②跑 -Mode collect 发采集包'
    }

    # ---- 5. 试登录 ----
    $cred = Get-ActiveCred
    if ($cred) {
        $a3 = (Read-Host '  现在试登录一次? [Y/n]').Trim().ToLowerInvariant()
        if ($a3 -ne 'n') {
            if (Test-Online) {
                Write-Stdout '  已经在线,跳过(要强制重登: -Mode login -Force)'
            } else {
                $ok = Invoke-LoginOnce -Cred $cred
                Write-Stdout ('  登录' + $(if ($ok) { '成功 ✓' } else { '失败 —— 看 -Mode log,或跑 -Mode portal -DetectPortal' }))
            }
        }
    }

    # ---- 6. 开机自启 ----
    if (-not (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) {
        $a4 = (Read-Host '  要装开机自启吗? [Y/n]').Trim().ToLowerInvariant()
        if ($a4 -ne 'n') { [void](Invoke-Install) }
    }

    Write-Stdout ''
    Write-Stdout '  向导结束。以后直接开菜单就能用。'
    Write-Stdout ''
}

# ================= 诊断 =================
function Add-DoctorItem {
    param($List, [string]$Name, [string]$Level, [string]$Detail)
    [void]$List.Add([pscustomobject]@{ Name = $Name; Level = $Level; Detail = $Detail })
}

function Invoke-Doctor {
    param([switch]$NoRepair)

    $items = New-Object System.Collections.ArrayList
    $lines = New-Object System.Collections.Generic.List[string]

    $lines.Add('===== 校园网自动认证 · 诊断报告 =====')
    $lines.Add('时间: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $lines.Add('脚本: ' + $SelfScript)
    $lines.Add('')

    # 1 无线
    try {
        $w = Get-WifiInfo
        if (-not $w.Interface) { Add-DoctorItem $items '无线网卡' 'WARN' '没有无线网卡(可能在用有线)' }
        elseif ($w.State -match '^(connected|已连接)$' -and $w.Ssid -eq $WifiSsid) { Add-DoctorItem $items '无线网卡' 'OK' ('已连接 ' + $w.Ssid) }
        elseif ($w.State -match '^(connected|已连接)$') { Add-DoctorItem $items '无线网卡' 'WARN' ('连的是 ' + $w.Ssid + ',目标 ' + $WifiSsid) }
        else { Add-DoctorItem $items '无线网卡' 'BAD' ('未连接(状态 ' + $w.State + ')') }
    } catch { Add-DoctorItem $items '无线网卡' 'BAD' ('读取失败: ' + $_.Exception.Message) }

    # 2 IP / 网关
    try {
        $cfg = @(Get-NetIPConfiguration -ErrorAction Stop | Where-Object { $_.IPv4DefaultGateway }) | Select-Object -First 1
        if ($cfg) {
            $ip = (@($cfg.IPv4Address)[0]).IPAddress
            Add-DoctorItem $items 'IP / 网关' 'OK' ($ip + ' / ' + $cfg.IPv4DefaultGateway.NextHop)
        } else {
            Add-DoctorItem $items 'IP / 网关' 'BAD' '没有默认网关'
        }
    } catch { Add-DoctorItem $items 'IP / 网关' 'WARN' ('读取失败: ' + $_.Exception.Message) }

    # 3 门户可达
    try {
        $code = & $Curl -s -o NUL -w '%{http_code}' --max-time 8 ('http://{0}/' -f $PortalHost) 2>$null
        if ($code -match '^[23]') { Add-DoctorItem $items '门户可达' 'OK' ('HTTP ' + $code) }
        else { Add-DoctorItem $items '门户可达' 'BAD' ('HTTP ' + $code) }
    } catch { Add-DoctorItem $items '门户可达' 'BAD' ('请求失败: ' + $_.Exception.Message) }

    # 3.1 代理环境(开着代理软件时 curl 会全部返回 000,是最容易误判成"校园网不通"的坑)
    if ($script:ProxyEnvHit -and $script:ProxyEnvHit.Count -gt 0) {
        Add-DoctorItem $items '代理环境变量' 'WARN' (($script:ProxyEnvHit -join ' , ') + ' —— 已在本进程内忽略(认证必须直连)')
    } else {
        Add-DoctorItem $items '代理环境变量' 'OK' '没有 http_proxy 之类干扰'
    }

    # 3.2 网关页面(很多校园网门户就架在网关上,门户域名不通时靠这个找)
    $gwIp = Get-DefaultGatewayIp
    if ($gwIp) {
        $gwCode = ''
        try { $gwCode = ((& $Curl -s -o NUL -A $UA -w '%{http_code}' --max-time 5 ('http://{0}/' -f $gwIp) 2>$null) -join '').Trim() } catch { }
        if ($gwCode -match '^[23]') { Add-DoctorItem $items '网关页面' 'OK' ($gwIp + ' 有响应 HTTP ' + $gwCode + '(门户可能就在这)') }
        else { Add-DoctorItem $items '网关页面' 'WARN' ($gwIp + ' HTTP ' + $gwCode) }
    } else {
        Add-DoctorItem $items '网关页面' 'WARN' '读不到默认网关'
    }

    # 4 DNS
    try {
        $dns = [System.Net.Dns]::GetHostAddresses('www.baidu.com') | Select-Object -First 1
        Add-DoctorItem $items 'DNS 解析' 'OK' ('www.baidu.com -> ' + $dns.IPAddressToString)
    } catch { Add-DoctorItem $items 'DNS 解析' 'BAD' ('解析失败: ' + $_.Exception.Message) }

    # 5 系统时间
    $srv = Get-PortalServerTime
    if ($srv) {
        $skew = ($srv.UtcDateTime - [DateTime]::UtcNow).TotalSeconds
        $lvl = 'OK'
        if ([Math]::Abs($skew) -gt 3) { $lvl = 'WARN' }
        Add-DoctorItem $items '系统时间' $lvl ('与门户相差 {0:N1} 秒' -f $skew)
    } else {
        Add-DoctorItem $items '系统时间' 'WARN' '取不到门户时间,无法比对'
    }

    # 6 凭据
    $store = Get-Store
    $masked = ($store.active -replace '(\d{3})\d{4}(\d{4})', '$1****$2')
    if ($store.accounts.Count -eq 0) {
        Add-DoctorItem $items '保存的账号' 'BAD' '还没有保存账号'
    } else {
        $cred = Get-ActiveCred
        if ($cred) { Add-DoctorItem $items '保存的账号' 'OK' ([string]$store.accounts.Count + ' 个,当前 ' + $masked + ',密码解密正常') }
        else { Add-DoctorItem $items '保存的账号' 'BAD' ('有 ' + $store.accounts.Count + ' 个,但当前账号解密失败') }
    }

    # 7 认证状态
    $ns = Get-NetStatus
    if ($ns.Online) { Add-DoctorItem $items '认证状态' 'OK' ('已通过认证 — ' + $ns.Detail) }
    elseif ($ns.Hijacked) { Add-DoctorItem $items '认证状态' 'BAD' ('未认证(被劫持: ' + $ns.Detail + ')') }
    else { Add-DoctorItem $items '认证状态' 'BAD' ('不通 — ' + $ns.Detail) }

    # 7.1 门户识别:在线时门户不会被劫持,别白探
    if ($ns.Online) {
        Add-DoctorItem $items '门户识别' 'OK' '已在线,门户没被劫持(要测识别先断网重连)'
    } else {
        $cap = Get-CaptiveRedirect -ProbeUrl @('http://www.msftconnecttest.com/connecttest.txt', 'http://connect.rom.miui.com/generate_204')
        if ($cap) {
            $fp = Get-PortalFingerprint $cap
            Add-DoctorItem $items '门户识别' 'OK' ($fp.Vendor + $(if ($fp.Id) { '(' + $fp.Id + ')' } else { '(无预设)' }) + ' ← ' + $cap)
        } else {
            Add-DoctorItem $items '门户识别' 'WARN' '探不到 302:门户可能只在 HTTPS 拦,或要手动开一次认证页'
        }
    }

    # 8 计划任务
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $task) {
        Add-DoctorItem $items '开机自启' 'WARN' '未安装计划任务'
    } else {
        $targ = ($task.Actions | Select-Object -First 1).Arguments
        if ($targ -match [regex]::Escape($WorkDir)) { Add-DoctorItem $items '开机自启' 'OK' ('已安装(' + $task.State + '),指向本目录') }
        else { Add-DoctorItem $items '开机自启' 'BAD' ('已安装,但指向别处: ' + $targ) }
    }

    # 9 进程
    $hasTray = [bool](Get-TrayProcess)
    $hasWatch = [bool](Get-WatcherProcess)
    if ($hasTray -and $hasWatch) { Add-DoctorItem $items '运行中的进程' 'OK' '托盘 + 守护都在' }
    elseif ($hasWatch) { Add-DoctorItem $items '运行中的进程' 'WARN' '只有守护,没有托盘' }
    elseif ($hasTray) { Add-DoctorItem $items '运行中的进程' 'WARN' '只有托盘,守护没起来' }
    else { Add-DoctorItem $items '运行中的进程' 'WARN' '都没在跑' }

    # 10 设置
    $set = Get-Settings
    Add-DoctorItem $items '设置' 'OK' ('WiFi 自动连接: ' + $(if ($set.autoWifi) { '开' } else { '关' }) + ',目标: ' + $set.wifiSsid)

    # 11 日志
    if (Test-Path $LogFile) {
        $tail = @(Get-Content $LogFile -Tail 200 -ErrorAction SilentlyContinue)
        $errs = @($tail | Where-Object { $_ -match '\[ERROR\]' })
        if ($errs.Count -gt 0) {
            $lastErr = ($errs | Select-Object -Last 1) -replace '^\S+ \S+ ', ''
            Add-DoctorItem $items '日志' 'WARN' ('最近 200 行有 ' + $errs.Count + ' 条 ERROR,最后: ' + $lastErr)
        } else {
            Add-DoctorItem $items '日志' 'OK' '最近 200 行没有 ERROR'
        }
    } else {
        Add-DoctorItem $items '日志' 'WARN' '还没有日志'
    }

    # 12 WiFi 列表(每次连接都会自动刷新)
    try {
        $wl = Update-WifiList -Force
        $vis = @($wl.Visible)
        if ($wl.TargetVisible -and $wl.TargetInProfiles) {
            Add-DoctorItem $items 'WiFi 列表' 'OK' ('可见 ' + $vis.Count + ' 个,目标可见(信号 ' + $wl.TargetSignal + '%)且有配置')
        } elseif ($wl.TargetVisible) {
            Add-DoctorItem $items 'WiFi 列表' 'WARN' ('可见 ' + $vis.Count + ' 个,目标可见但 Windows 里没有配置')
        } else {
            Add-DoctorItem $items 'WiFi 列表' 'WARN' ('可见 ' + $vis.Count + ' 个,列表里没有 ' + $WifiSsid)
        }
    } catch { Add-DoctorItem $items 'WiFi 列表' 'WARN' ('读取失败: ' + $_.Exception.Message) }

    # 13 网速采样
    try {
        [void](Get-NetSpeed -Reset)
        [void](Get-NetSpeed)
        Start-Sleep -Milliseconds 700
        $spd = Get-NetSpeed
        if ($spd -and $spd.Ready) {
            Add-DoctorItem $items '网速采样' 'OK' ('当前 ↓ ' + $spd.DownText + ' ↑ ' + $spd.UpText + ',网卡计数器读取正常')
        } else {
            Add-DoctorItem $items '网速采样' 'WARN' '读不到网卡字节计数'
        }
    } catch { Add-DoctorItem $items '网速采样' 'WARN' ('失败: ' + $_.Exception.Message) }

    # 14 自动修复:只在诊断出「常见毛病」时动手,修完复查
    if (-not $NoRepair) {
        $fixable = @($items | Where-Object { $_.Level -eq 'BAD' -and @('无线网卡', '门户可达', 'DNS 解析', '认证状态') -contains $_.Name })
        if ($fixable.Count -gt 0) {
            $names = @($fixable | ForEach-Object { $_.Name }) -join '/'
            try {
                $rep = Invoke-Repair -Level 'safe' -Auto -Why ('诊断发现 ' + $names)
                if ($rep.Ran) {
                    $doneNames = @($rep.Steps | Where-Object { $_.Ok } | ForEach-Object { $_.Name })
                    $failNames = @($rep.Steps | Where-Object { -not $_.Ok } | ForEach-Object { $_.Name })
                    $txt = '已尝试 ' + $(if ($doneNames.Count) { $doneNames -join '、' } else { '(无)' })
                    if ($failNames.Count) { $txt += ';没成功:' + ($failNames -join '、') }
                    $txt += $(if ($rep.Online) { ' → 现在已在线' } else { ' → 仍不通' })
                    Add-DoctorItem $items '自动修复' $(if ($rep.Online) { 'OK' } else { 'WARN' }) $txt
                } elseif ($rep.Note) {
                    Add-DoctorItem $items '自动修复' 'WARN' $rep.Note
                }
            } catch {
                Add-DoctorItem $items '自动修复' 'WARN' ('执行失败: ' + $_.Exception.Message)
            }
        }
    }

    # 打印
    $i = 0; $ok = 0; $warn = 0; $bad = 0
    foreach ($it in $items) {
        $i++
        if ($it.Level -eq 'OK') { $mark = '[OK]'; $ok++ }
        elseif ($it.Level -eq 'WARN') { $mark = '[!] '; $warn++ }
        else { $mark = '[X] '; $bad++ }
        $lines.Add(('{0,2}. {1} {2}  {3}' -f $i, $mark, $it.Name, $it.Detail))
    }
    $lines.Add('')
    $lines.Add(('汇总: {0} 项正常,{1} 项需注意,{2} 项异常' -f $ok, $warn, $bad))

    $report = $lines -join "`r`n"
    Write-Stdout $report

    $reportPath = Join-Path $WorkDir 'doctor-report.txt'
    try {
        Set-Content -Path $reportPath -Value $report -Encoding UTF8
        Write-Stdout ('报告已保存: ' + $reportPath)
    } catch { }

    if ($bad -gt 0) { return $false }
    return $true
}

# ================= 门户改版自检 =================
function Invoke-PortalCheck {
    param([switch]$Deep)

    $pf0 = Get-PortalProfile
    if ([string]$pf0.protocol -eq 'esurfing') {
        # 天翼没有 Dr.COM 那套页面 + RC4 指纹,自检换成「接口是否还认这套协议」
        Write-Log '门户自检:当前是天翼协议,改为探测 enet.10000.gd.cn 接口' 
        $d = Invoke-EsDetect
        return $true
    }

    $items = New-Object System.Collections.ArrayList
    $lines = New-Object System.Collections.Generic.List[string]
    $pageUrl = 'http://{0}{1}' -f $PortalHost, $PortalPage

    $lines.Add('===== 门户改版自检 =====')
    $lines.Add('时间: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $lines.Add('门户页: ' + $pageUrl)
    $lines.Add('')

    # A. 页面能否打开(一次请求:body 落临时文件,同时拿状态码和最终地址)
    $html = ''
    $fcode = ''
    $feff = ''
    $tmpFile = Join-Path ([System.IO.Path]::GetTempPath()) ('pc-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.html')
    try {
        $meta = & $Curl -s -L -o $tmpFile -w '%{http_code}|%{url_effective}' --max-time 20 $pageUrl 2>$null
        $parts = ($meta -join '') -split '\|'
        if ($parts.Count -ge 1) { $fcode = $parts[0] }
        if ($parts.Count -ge 2) { $feff = $parts[1] }
        if (Test-Path $tmpFile) { $html = [System.IO.File]::ReadAllText($tmpFile) }
    } catch { } finally {
        Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
    }

    if ([string]::IsNullOrWhiteSpace($html)) {
        Add-DoctorItem $items '门户页可访问' 'BAD' '取不到页面 —— 模板 ID 可能变了,用浏览器地址栏的完整 URL 更新 -PortalPage'
    } elseif ($fcode -notmatch '^2') {
        Add-DoctorItem $items '门户页可访问' 'BAD' ('HTTP ' + $fcode + ',页面打不开')
    } elseif (($feff -match 'ac_portal') -or ($html -match '(认证|ac_portal|template=)')) {
        # 门户页里链接都是 "../share/..." 相对路径,正文不一定出现 ac_portal 字面量
        $detail = 'HTTP ' + $fcode + ',页面 ' + $html.Length + ' 字符'
        if ($feff -match 'ac_portal') { $detail = $detail + ',最终地址在门户内' }
        Add-DoctorItem $items '门户页可访问' 'OK' $detail
    } else {
        Add-DoctorItem $items '门户页可访问' 'WARN' '页面能打开,但没看到门户特征(可能是别的页面)'
    }

    # B. 抓前端脚本做指纹比对
    $markers = @(
        @{ Name = '登录接口 login.php'; Pattern = 'login\.php' },
        @{ Name = '密码登录参数 pwdLogin'; Pattern = 'pwdLogin' },
        @{ Name = 'RC4 加密 do_encrypt_rc4'; Pattern = 'do_encrypt_rc4' },
        @{ Name = '时间戳参数 auth_tag'; Pattern = 'auth_tag' }
    )

    $jsText = ''
    if ($html) {
        $srcs = @([regex]::Matches($html, 'src=["'']([^"'']+\.js[^"'']*)["'']') | ForEach-Object { $_.Groups[1].Value })
        foreach ($src in @($srcs | Where-Object { $_ -match '(logic_new|pcauth)\.js' })) {
            try {
                $abs = ([System.Uri]::new([System.Uri]$pageUrl, $src)).AbsoluteUri
                $jsText += ((& $Curl -s -L --max-time 20 $abs 2>$null) -join "`n")
            } catch { }
        }
    }

    if ([string]::IsNullOrWhiteSpace($jsText)) {
        Add-DoctorItem $items '前端脚本' 'WARN' '没抓到 logic_new.js / pcauth.js,无法做指纹比对'
    } else {
        Add-DoctorItem $items '前端脚本' 'OK' ('已抓取 ' + $jsText.Length + ' 字符')
        foreach ($m in $markers) {
            if ($jsText -match $m.Pattern) { Add-DoctorItem $items $m.Name 'OK' '存在' }
            else { Add-DoctorItem $items $m.Name 'BAD' '没找到 —— 门户可能改版,需要重新逆向' }
        }
    }

    # C. 接口活体探测(只在 -Deep 时做,会用假账号发一次请求)
    if ($Deep) {
        try {
            $tag = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds().ToString()
            $body = 'opr=pwdLogin&userName=00000000000&pwd=' + (Get-Rc4Hex -Text 'x' -Key $tag) + '&auth_tag=' + $tag + '&rememberPwd=0'
            $raw = ((& $Curl -s -X POST ('http://{0}/ac_portal/login.php' -f $PortalHost) `
                        -H 'Content-Type: application/x-www-form-urlencoded' `
                        -H ('User-Agent: {0}' -f $UA) `
                        --data $body --max-time 20 2>$null) -join '')
            $j = $null
            try { $j = $raw | ConvertFrom-Json } catch { }
            if (-not $j) {
                Add-DoctorItem $items '接口活体探测' 'BAD' ('返回的不是 JSON,原文: ' + $raw.Trim())
            } elseif ($null -eq $j.success) {
                Add-DoctorItem $items '接口活体探测' 'BAD' ('返回里没有 success 字段,原文: ' + $raw.Trim())
            } elseif ($j.msg -match '用户名或密码错误') {
                Add-DoctorItem $items '接口活体探测' 'OK' '接口正常(假账号被正确拒绝)'
            } else {
                Add-DoctorItem $items '接口活体探测' 'WARN' ('有响应但措辞变了: ' + $j.msg)
            }
        } catch {
            Add-DoctorItem $items '接口活体探测' 'WARN' ('探测失败: ' + $_.Exception.Message)
        }
    } else {
        Add-DoctorItem $items '接口活体探测' 'OK' '已跳过(加 -Deep 才会真的发一次假账号请求)'
    }

    # 输出
    $i = 0; $ok = 0; $warn = 0; $bad = 0
    foreach ($it in $items) {
        $i++
        if ($it.Level -eq 'OK') { $mark = '[OK]'; $ok++ }
        elseif ($it.Level -eq 'WARN') { $mark = '[!] '; $warn++ }
        else { $mark = '[X] '; $bad++ }
        $lines.Add(('{0,2}. {1} {2}  {3}' -f $i, $mark, $it.Name, $it.Detail))
    }
    $lines.Add('')
    if ($bad -eq 0) { $lines.Add('结论: 指纹全部匹配,门户未改版') }
    else { $lines.Add('结论: 检测到 ' + $bad + ' 处变化,脚本可能需要更新(把这份报告发给维护者)') }
    $lines.Add(('汇总: {0} 项正常,{1} 项需注意,{2} 项变化' -f $ok, $warn, $bad))

    $report = $lines -join "`r`n"
    Write-Stdout $report

    $rp = Join-Path $WorkDir 'portal-check.txt'
    try {
        Set-Content -Path $rp -Value $report -Encoding UTF8
        Write-Stdout ('报告已保存: ' + $rp)
    } catch { }

    # 记录本次自检时间(守护靠它决定 24 小时内不再重复)
    try {
        $s = Get-Settings
        $s.lastPortalCheck = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Save-Settings -Settings $s
    } catch { }

    # 写一行日志,后台自动跑的时候也能在 campus-net.log 里看到结果
    if ($bad -gt 0) {
        Write-Log ('门户自检: 检测到 ' + $bad + ' 处变化,门户可能改版(详见 portal-check.txt)') 'WARN'
        return $false
    }
    Write-Log '门户自检通过: 登录接口 / 参数 / RC4 / auth_tag 四处指纹都在'
    return $true
}

# ================= 窗口背景图(Windows Terminal) =================
function Get-WTSettingsPath {
    $cands = @(
        (Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json'),
        (Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows Terminal\settings.json'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows Terminal Preview\settings.json')
    )
    foreach ($c in $cands) { if (Test-Path $c) { return $c } }
    return $null
}

function Set-TerminalBackground {
    param([string]$ImagePath, [switch]$Clear, [double]$Opacity = 0.35)

    $cfg = Get-WTSettingsPath
    if (-not $cfg) {
        Write-Log '没找到 Windows Terminal 设置文件 —— 控制台(conhost)本身不支持背景图,需要装 Windows Terminal' 'WARN'
        return $false
    }

    if (-not $Clear) {
        if (-not $ImagePath) { Write-Log '没有指定图片路径' 'ERROR'; return $false }
        if (-not (Test-Path $ImagePath)) { Write-Log ('图片不存在: ' + $ImagePath) 'ERROR'; return $false }
        $ext = [System.IO.Path]::GetExtension($ImagePath).ToLower()
        if (-not (@('.png', '.jpg', '.jpeg', '.bmp', '.gif') -contains $ext)) {
            Write-Log ('不支持的图片格式: ' + $ext + '(支持 png / jpg / bmp / gif)') 'ERROR'
            return $false
        }
        $ImagePath = (Resolve-Path $ImagePath).Path
    }

    # 改之前先备份
    $backup = $cfg + '.bak-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    Copy-Item $cfg $backup -Force
    Write-Log ('已备份原设置: ' + $backup)

    try {
        $json = Get-Content $cfg -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log ('解析 Windows Terminal 设置失败(可能被手动加过注释):' + $_.Exception.Message) 'ERROR'
        return $false
    }

    # 目标是默认 profile(双击 connect.bat 打开的就是它)
    $defGuid = $json.defaultProfile
    $target = @(@($json.profiles.list) | Where-Object { $_.guid -eq $defGuid })[0]
    if (-not $target) { $target = @(@($json.profiles.list) | Where-Object { $_.name -match 'PowerShell' })[0] }
    if (-not $target) { Write-Log '没找到默认 profile,为避免改坏你的终端,已放弃' 'ERROR'; return $false }

    if ($Clear) {
        foreach ($k in 'backgroundImage', 'backgroundImageOpacity', 'backgroundImageStretchMode') {
            if ($target.PSObject.Properties[$k]) { $target.PSObject.Properties.Remove($k) }
        }
        Write-Log ('已清除「' + $target.name + '」的背景图') 'OK'
    } else {
        $target | Add-Member -NotePropertyName backgroundImage -NotePropertyValue $ImagePath -Force
        $target | Add-Member -NotePropertyName backgroundImageOpacity -NotePropertyValue $Opacity -Force
        $target | Add-Member -NotePropertyName backgroundImageStretchMode -NotePropertyValue 'uniformToFill' -Force
        Write-Log ('已给「' + $target.name + '」设置背景图: ' + $ImagePath + ' (透明度 ' + $Opacity + ')') 'OK'
    }

    $out = $json | ConvertTo-Json -Depth 20
    # 写回之前先确认是合法 JSON,避免把终端配置写坏
    try { $null = $out | ConvertFrom-Json } catch { Write-Log '生成的新设置不是合法 JSON,已放弃写入(原文件未改动)' 'ERROR'; return $false }

    # 必须写成「无 BOM 的 UTF-8」,带 BOM 会让 Windows Terminal 解析失败
    [System.IO.File]::WriteAllText($cfg, $out, (New-Object System.Text.UTF8Encoding($false)))
    Write-Log '设置已更新 —— 关掉当前窗口重开(或新建标签页)才看得到效果'
    return $true
}

# ================= 退出程序时的三个选择 =================
function Set-ProgramWindowMinimized {
    # 「最小化程序」:窗口收起来,进程照跑。
    # Windows Terminal 里控制台窗口本身是隐藏的,这时退一步把前台窗口(基本就是那个终端)最小化。
    if (-not ('CampusNetWin32' -as [type])) {
        try {
            $cs = @'
using System;
using System.Runtime.InteropServices;
public class CampusNetWin32 {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
'@
            Add-Type -TypeDefinition $cs -Language CSharp
        } catch { }
    }
    if (-not ('CampusNetWin32' -as [type])) { return $false }
    $h = [CampusNetWin32]::GetConsoleWindow()
    if ($h -ne [IntPtr]::Zero -and [CampusNetWin32]::IsWindowVisible($h)) { return [CampusNetWin32]::ShowWindow($h, 6) }
    $fg = [CampusNetWin32]::GetForegroundWindow()
    if ($fg -ne [IntPtr]::Zero) { return [CampusNetWin32]::ShowWindow($fg, 6) }
    return $false
}

function Show-ExitChooser {
    # 关程序窗口前问一句:完全退出 / 最小化
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '退出校园网自动登录'
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.TopMost = $true
    $f.ClientSize = [System.Drawing.Size]::new(472, 214)

    $lb = New-Object System.Windows.Forms.Label
    $lb.Text = '程序窗口要关掉了,你想留下什么?'
    $lb.Location = [System.Drawing.Point]::new(18, 14)
    $lb.Size = [System.Drawing.Size]::new(436, 26)
    try { $lb.Font = [System.Drawing.Font]::new('Microsoft YaHei UI', 10.5, [System.Drawing.FontStyle]::Bold) } catch { }
    $f.Controls.Add($lb)

    $script:ExitChoice = 'cancel'
    foreach ($o in @(
            @{ t = '完全退出程序 —— 托盘、守护全停'; v = 'quit'; y = 54 },
            @{ t = '最小化程序 —— 窗口收起来,一切照常跑'; v = 'min'; y = 118 }
        )) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = [string]$o.t
        $b.Location = [System.Drawing.Point]::new(18, [int]$o.y)
        $b.Size = [System.Drawing.Size]::new(436, 44)
        $b.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $b.Tag = [string]$o.v
        $b.add_Click({
                param($snd, $e)
                $script:ExitChoice = [string]$snd.Tag
                $snd.FindForm().Close()
            })
        if ($o.v -eq 'min') { $f.AcceptButton = $b }
        $f.Controls.Add($b)
    }

    $bCancel = New-Object System.Windows.Forms.Button
    $bCancel.Text = '取消(Esc) —— 回到菜单'
    $bCancel.Location = [System.Drawing.Point]::new(18, 168)
    $bCancel.Size = [System.Drawing.Size]::new(250, 36)
    $bCancel.Tag = 'cancel'
    $bCancel.add_Click({
            param($snd, $e)
            $script:ExitChoice = 'cancel'
            $snd.FindForm().Close()
        })
    $f.CancelButton = $bCancel
    $f.Controls.Add($bCancel)

    [void]$f.ShowDialog()
    $r = $script:ExitChoice
    try { $f.Dispose() } catch { }
    return $r
}

function Invoke-ExitChoice {
    # 返回 'exit' = 真的退出脚本;'stay' = 回到菜单继续跑
    param([string]$Choice)
    switch ($Choice) {
        'quit' {
            Stop-Watcher
            [void](Stop-Tray)
            Write-Stdout ''
            Write-Stdout '已完全退出:守护、托盘都停了。'
            Write-Stdout '开机自启还在,下次登录 Windows 会自己起来(不想要就选 10 卸载开机自启)。'
            Write-Log '用户选择:完全退出程序' 'OK'
            return 'exit'
        }
        'min' {
            Write-Stdout ''
            if (Set-ProgramWindowMinimized) {
                Write-Stdout '窗口已最小化:程序继续在后台跑,从任务栏点回来就行'
            } else {
                Write-Stdout '没法自动最小化,请手动点窗口右上角的最小化;程序会继续运行'
            }
            return 'stay'
        }
        default {
            Write-Stdout '已取消退出'
            return 'stay'
        }
    }
}

# ================= 菜单 =================
function Show-Menu {
    while ($true) {
        $store = Get-Store
        $installed = [bool](Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)
        $st = Read-State

        Write-Stdout ''
        Write-Stdout '=================================================='
        Write-Stdout ('            校园网自动登录  v' + $script:AppVersion + ' ' + $script:AppEdition)
        Write-Stdout '            (仅校园网 + 内存清理,不含桌面宠物)'
        Write-Stdout '=================================================='
        if ($store.accounts.Count -eq 0) {
            Write-Stdout '  当前账号: (未设置)'
        } else {
            Write-Stdout ('  当前账号: ' + $store.active + '   [已保存 ' + $store.accounts.Count + ' 个]')
        }
        if ($installed) { Write-Stdout '  开机自启: 已安装' } else { Write-Stdout '  开机自启: 未安装' }
        $curWifi = Get-WifiInfo
        $wl = Update-WifiList
        $visList = @($wl.Visible)
        Write-Stdout ('  WiFi 自动连接: ' + $(if ($script:AutoWifi) { '开' } else { '关' }) + '   目标: ' + $WifiSsid)
        Write-Stdout ('  当前 WiFi: ' + $(if ($curWifi.Ssid) { $curWifi.Ssid } else { '(未连接)' }) + '   可见 ' + $visList.Count + ' 个网络')
        $sp = Get-NetSpeed
        if (-not $sp.Ready) { Start-Sleep -Milliseconds 600; $sp = Get-NetSpeed }
        Write-Stdout ('  实时网速: ↓ ' + $sp.DownText + '   ↑ ' + $sp.UpText + '   峰值 ↓ ' + $sp.PeakDownText)
        if ($st) {
            $netTxt = $(if ($st.online) { '在线' } else { '离线' })
            Write-Stdout ('  最近状态: ' + $netTxt + '  (' + $st.timestamp + ')')
        } else {
            Write-Stdout '  最近状态: —'
        }
        Write-Stdout '--------------------------------------------------'
        Write-Stdout '  1) WiFi 自动连接: 开/关切换(优先级最高,启动即联网)'
        Write-Stdout '  2) 修改目标 WiFi 名称'
        Write-Stdout '  3) 添加账号 / 更换密码'
        Write-Stdout '  4) 切换当前账号'
        Write-Stdout '  5) 删除账号'
        Write-Stdout '  6) 查看已保存的账号'
        Write-Stdout '  7) 立刻登录一次'
        Write-Stdout '  8) 校正系统时间'
        Write-Stdout '  9) 安装开机自启'
        Write-Stdout '  10) 卸载开机自启'
        Write-Stdout '  B) 修改窗口背景图'
        Write-Stdout '  D) 一键诊断(出报告)'
        Write-Stdout '  R) 一键修复(重置网卡 / 刷 DNS / 清 Cookie)'
        Write-Stdout '  O) 门户类型 / 通用认证配置'
        $curProto = [string](Get-PortalProfile).name
        Write-Stdout ('       (当前: ' + $curProto + ')')
        Write-Stdout '  C) 采集诊断包(远程排查用,自动复制到剪贴板)'
        Write-Stdout '  G) 新手向导(问一句是不是校园网 + 认门户)'
        Write-Stdout '  L) 查看日志'
        Write-Stdout '  0) 退出(可选:完全退出 / 最小化)'
        Write-Stdout '=================================================='
        $c = (Read-Host '请选择').Trim().ToUpper()

        switch ($c) {
            '1' {
                $s1 = Get-Settings
                $s1.autoWifi = -not [bool]$s1.autoWifi
                Save-Settings -Settings $s1
                $script:AutoWifi = [bool]$s1.autoWifi
                Write-Log ('WiFi 自动连接已' + $(if ($script:AutoWifi) { '开启' } else { '关闭' })) 'OK'
                if ($script:AutoWifi) { [void](Connect-CampusWifi -Ssid $WifiSsid) }
            }
            '2' {
                Write-Stdout ('当前目标 WiFi: ' + $WifiSsid)
                $n = Read-Host '输入新的 WiFi 名称(回车取消)'
                if (-not [string]::IsNullOrWhiteSpace($n)) {
                    $s2 = Get-Settings
                    $s2.wifiSsid = $n.Trim()
                    Save-Settings -Settings $s2
                    $script:WifiSsid = $s2.wifiSsid
                    Write-Log ('目标 WiFi 已改为: ' + $s2.wifiSsid) 'OK'
                }
            }
            '3' { Invoke-Setup -IsChange }
            '4' { Invoke-Switch }
            '5' { Invoke-Delete }
            '6' { Show-Accounts }
            '7' {
                $cred = Resolve-Cred -AllowPrompt
                if ($cred) { [void](Invoke-LoginOnce -Cred $cred) } else { Write-Stdout '没有账号,先选 3 添加' }
            }
            '8' { [void](Sync-Clock -Always) }
            '9' { [void](Invoke-Install) }
            '10' { Invoke-Uninstall }
            'B' {
                $termNow = $(if ($env:WT_SESSION) { 'Windows Terminal' } else { '当前控制台不是 Windows Terminal(conhost 不支持背景图)' })
                Write-Stdout ('  终端: ' + $termNow)
                $img = Read-Host '图片完整路径(把图片拖进来即可;输入 clear 清除背景;回车取消)'
                if (-not [string]::IsNullOrWhiteSpace($img)) {
                    $img = $img.Trim().Trim('"').Trim("'")
                    if ($img -eq 'clear') { [void](Set-TerminalBackground -Clear) }
                    else { [void](Set-TerminalBackground -ImagePath $img) }
                }
            }
            'D' { [void](Invoke-Doctor) }
            'R' { [void](Invoke-Repair -Level 'full' -Force) }
            'O' {
                Show-PortalProfile
                $pid2 = Read-Host '输入预设 id 切换 / detect 探测 / auto 自动适配(回车取消)'
                if (-not [string]::IsNullOrWhiteSpace($pid2)) {
                    $pid2 = $pid2.Trim()
                    if ($pid2 -eq 'detect') { Invoke-PortalDetect }
                    elseif ($pid2 -eq 'auto') {
                        $ra = Invoke-PortalAdapt
                        if ($ra.Changed) { Write-Stdout ('已自动适配: ' + $ra.Vendor + '(' + $ra.Reason + ')') }
                        else { Write-Stdout ('没有变化: ' + $ra.Reason) }
                    }
                    else { [void](Set-PortalProfile -Id $pid2) }
                }
            }
            'L' { Show-Log }
            'C' { Invoke-Collect }
            'G' { Invoke-Wizard }
            '0' {
                $script:MenuExit = Invoke-ExitChoice -Choice (Show-ExitChooser)
                if ($script:MenuExit -eq 'exit') { return }
            }
            default { Write-Stdout '无效选择' }
        }
    }
}

# ================= 主流程 =================
if (-not $Curl) {
    Write-Log '找不到 curl.exe(Windows 10 1803+ 自带),请确认系统版本' 'ERROR'
    exit 2
}

# ===== 代理:校园网认证必须直连 =====
# 开着代理软件时,curl 会跟着 http_proxy 走,对校园网地址一律 000 ——
# 表现就是「门户可达 HTTP 000 / 探测返回 HTTP 000」这种查无实据的失败。
$script:ProxyEnvHit = @()
foreach ($pn in @('HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY')) {
    $pv = [Environment]::GetEnvironmentVariable($pn)
    if (-not [string]::IsNullOrWhiteSpace($pv)) {
        $script:ProxyEnvHit += ($pn + '=' + $pv)
        [Environment]::SetEnvironmentVariable($pn, $null)
    }
}
if ($script:ProxyEnvHit.Count -gt 0) {
    Write-Log ('检测到代理环境变量,已在本进程内忽略(校园网认证必须直连):' + ($script:ProxyEnvHit -join ' , ')) 'WARN'
}

# ===== 门户配置:命令行没显式指定 -PortalHost/-PortalPage 时,以 portal.json 为准 =====
$profileNow = Get-PortalProfile
if (-not $PSBoundParameters.ContainsKey('PortalHost') -and $profileNow.host) { $PortalHost = [string]$profileNow.host }
if (-not $PSBoundParameters.ContainsKey('PortalPage') -and $profileNow.page) { $PortalPage = [string]$profileNow.page }
if ($ProbeUrl -and $ProbeUrl.Count -gt 0) { $script:CaptiveProbes = $ProbeUrl }

# ===== 最高优先级:程序一启动就先连 WiFi,再做别的事 =====
if (@('menu', 'tray', 'watch', 'login', 'status', 'wifi') -contains $Mode) {
    if ($Mode -eq 'wifi') {
        # 用户明确要求连 WiFi:不看开关,直接连
        [void](Connect-CampusWifi -Ssid $WifiSsid)
    } elseif (-not $NoWifi -and $script:AutoWifi) {
        [void](Connect-CampusWifi -Ssid $WifiSsid)
    }
}

switch ($Mode) {
    'setup'     { Invoke-Setup }
    'switch'    { Invoke-Switch }
    'delete'    { Invoke-Delete }
    'list'      { Show-Accounts }
    'install'   { if (-not (Invoke-Install)) { exit 1 } }
    'uninstall' { Invoke-Uninstall }
    'log'       { Show-Log }
    'collect'   { Invoke-Collect; exit 0 }
    'wizard'    { Invoke-Wizard; exit 0 }

    'synctime'  { [void](Sync-Clock -Always) }

    'wifi'      { if (Test-WifiConnected -Ssid $WifiSsid) { exit 0 } else { exit 1 } }

    'doctor'    { if (Invoke-Doctor -NoRepair:$NoRepair) { exit 0 } else { exit 1 } }

    'repair' {
        $r = Invoke-Repair -Level $RepairLevel -Force -Why '手动一键修复'
        if ($r.Ran) {
            $doneNames = @($r.Steps | Where-Object { $_.Ok } | ForEach-Object { $_.Name })
            $failNames = @($r.Steps | Where-Object { -not $_.Ok } | ForEach-Object { $_.Name })
            $txt = $(if ($doneNames.Count) { $doneNames -join '、' } else { '没有可做的' })
            if ($failNames.Count) { $txt += ';没成功:' + ($failNames -join '、') }
            $txt += $(if ($r.Online) { ' → 已恢复在线' } else { ' → 还是不通' })
            [void](Send-Notify -Title $(if ($r.Online) { '修复完成' } else { '修复没成功' }) -Message $txt `
                    -Kind $(if ($r.Online) { 'info' } else { 'warn' }) -Tag 'repairdone' -Actions @('查看日志|campusnet:log'))
        }
        if ($r.Online) { exit 0 } else { exit 1 }
    }

    'nicreset' {
        # 提权助手任务专用:只做「重置网卡 + 重连 WiFi」,被自动修复调用
        $r = Repair-Nic -Name $NicName
        if ($r.Ok -and -not $NoWifi) { [void](Repair-Wifi -Ssid $WifiSsid) }
        if ($r.Ok) { exit 0 }
        Write-Log ('提权重置网卡失败:' + $r.Detail) 'ERROR'
        exit 1
    }
    'memclean' {
        # 提权助手任务专用:被托盘「系统清理」唤起,清 standby + 文件缓存
        $r = Invoke-MemoryClean -Scope $Scope
        $okAll = $true
        foreach ($p in $r.Parts) { if (-not $p.Ok) { $okAll = $false } }
        if ($okAll) { exit 0 } else { exit 1 }
    }

    'notify' {        # 手动验一下通知到底通不通(排查"我这边没弹通知"用)
        [void](Register-AppIdentity -Force)
        $a = Send-Notify -Title '校园网自动登录' -Message '通知中心适配自检:能看到这条就说明通了' -Actions @('打开菜单|campusnet:menu', '立即登录|campusnet:login')
        Write-Stdout ('通知身份: ' + $script:ToastAumid)
        Write-Stdout ('发送结果: ' + $(if ($a) { '已提交给通知中心' } else { '失败(只能写日志)' }))
        Write-Stdout ('开始菜单快捷方式: ' + (Get-StartMenuLinkPath))
        Write-Stdout ('快捷方式存在: ' + (Test-Path (Get-StartMenuLinkPath)))
        try {
            $h = [Windows.UI.Notifications.ToastNotificationManager]::History.GetHistory($script:ToastAumid)
            Write-Stdout ('通知中心里本应用的历史条数: ' + $h.Count)
        } catch { }
        if ($a) { exit 0 } else { exit 1 }
    }

    'check'     { if (Invoke-PortalCheck -Deep:$Deep) { exit 0 } else { exit 1 } }

    'portal' {
        if ($PortalId) {
            if (-not (Set-PortalProfile -Id $PortalId)) {
                Write-Stdout ('没有这个预设: ' + $PortalId)
                Show-PortalProfile
                exit 1
            }
        }
        if ($DetectPortal) { Invoke-PortalDetect; exit 0 }
        if ($AutoPortal) {
            # 看现在被劫持到哪,顺手把 portal.json 改对
            $r = Invoke-PortalAdapt
            if ($r.Changed) { Write-Stdout ('已自动适配: ' + $r.Vendor + '(' + $r.Reason + ')') }
            else { Write-Stdout ('没有变化: ' + $r.Reason + $(if ($r.Url) { '  门户: ' + $r.Url } else { '' })) }
            Show-PortalProfile
            exit 0
        }
        Show-PortalProfile
        exit 0
    }

    'background' {
        if ($ClearBackground) {
            if (Set-TerminalBackground -Clear) { exit 0 } else { exit 1 }
        } elseif ($BackgroundImage) {
            if (Set-TerminalBackground -ImagePath $BackgroundImage -Opacity $BackgroundOpacity) { exit 0 } else { exit 1 }
        } else {
            Write-Log '用法: -Mode background -BackgroundImage <图片路径> [-BackgroundOpacity 0.3] / 或 -ClearBackground' 'ERROR'
            exit 1
        }
    }

    'status' {
        $online = Test-Online
        $cred = Get-ActiveCred
        $acc = $(if ($cred) { $cred.User } else { '—' })
        if ($online) { Write-Log ('当前在线  账号:' + $acc) 'OK'; exit 0 }
        else { Write-Log ('当前离线(需要认证)  账号:' + $acc) 'WARN'; exit 1 }
    }

    'login' {
        $cred = Resolve-Cred -Account $User -PlainPassword $Password -AllowPrompt:(-not $Background)
        if (-not $cred) { Write-Log '没有可用凭据(后台模式不会弹输入框)' 'ERROR'; exit 1 }
        if ((Test-Online) -and -not $Force) { Write-Log '已经在线,无需登录(要强制重登加 -Force)' 'OK'; exit 0 }
        if (-not (Test-Online) -and -not $NoWifi) { [void](Ensure-Wifi -Ssid $WifiSsid) }
        $loginOk = Invoke-LoginOnce -Cred $cred
        if (-not $loginOk -and $Resume) {
            [void](Send-Notify -Title '校园网重认证失败' -Message '唤醒后连不上,守护会继续重试并自动修复' -Kind 'warn' `
                    -Tag 'resumefail' -MinIntervalSec 120 -Actions @('一键修复|campusnet:repair'))
        }
        if ($loginOk) { exit 0 } else { exit 1 }
    }

    'watch' { Invoke-Watch }

    'tray'  { Invoke-Tray }

    default {
        # 第一次跑(既没设置也没账号)先走一遍向导:问一句"这个 WiFi 是不是校园网",再认门户
        if ((-not (Test-Path $SettingsFn)) -and (-not (Test-Path $AccountsFn)) -and (-not (Test-Path $script:PortalFn))) {
            Invoke-Wizard
        }
        Show-Menu
    # 选过「完全退出」就别再把托盘拉回来了
        if ($script:MenuExit -eq 'exit') { exit 0 }
        # 菜单关掉之后:装了自启就把托盘拉回来;没装就明确告诉用户为啥没有常驻图标
        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            [void](Start-TrayIfNeeded)
        } else {
            Write-Stdout ''
        Write-Stdout '提示: 还没装「开机自启」,所以关掉这个窗口后不会有托盘图标常驻。'
            Write-Stdout '      菜单里选 9) 安装开机自启 就会一直留在右下角。'
        }
    }
}
