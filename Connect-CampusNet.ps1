#Requires -Version 5.1
<#
.SYNOPSIS
    校园网自动登录 —— 菜单 / 托盘 / 开机自启 / 后台守护 / 多账号
    支持中国电信 Dr.COM、广东电信天翼校园网,以及可配置的通用表单门户

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
    wifi | doctor | check | background | log | repair | notify | nicreset | pet | weather | portal

.NOTES
    桌面宠物的动画素材来自开源项目 dsh-pet(MIT License):
        https://github.com/PC2005-cloud/dsh-pet
    本脚本只用了它 assets/webm 里的动画,离线转成 PNG 帧后本地播放;
    其中不含该项目的任何代码或运行时。素材版权归原作者,详见 pet/NOTICE.txt。
#>
[CmdletBinding()]
param(
    [ValidateSet('menu', 'setup', 'switch', 'delete', 'list', 'login', 'status', 'watch', 'tray', 'install', 'uninstall', 'synctime', 'wifi', 'doctor', 'check', 'background', 'log', 'repair', 'notify', 'nicreset', 'pet', 'weather', 'portal', 'collect', 'wizard')]
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

    # 关掉桌面宠物(-Mode pet -PetOff)
    [switch]$PetOff,

    # 手动指定网卡(一般不用给,默认自动找无线网卡)
    [string]$NicName,

    # 门户配置(通用认证)
    [string]$PortalId,
    [switch]$DetectPortal,
    [switch]$AutoPortal,

    # 天气:彩云天气 token(不填就用 Open-Meteo 兜底)
    [string]$CaiyunToken,

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
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
$script:AppVersion = '1.2.0'
$Curl = (Get-Command curl.exe -ErrorAction SilentlyContinue).Source
$script:NicList = $null
$script:NicListAt = [datetime]::MinValue

# ---------------- 设置(可保存的开关) ----------------
function Get-Settings {
    $def = [pscustomobject]@{
        autoWifi        = $true
        wifiSsid        = 'CIT-Chinanet'
        lastPortalCheck = ''
        lastRepair      = ''
        lastNicReset    = ''
        lastResume      = ''
        pet             = $false
        petX            = 0
        petY            = 0
        petRandomMin    = 5
        petWander       = $true
        petWeather      = $true
        petScale        = 1.0
        petGravity      = 1400
        petRestitution  = 0.78
        petGroundFriction = 2.5
        petThrowPower   = 1.0
        petCeilingBounce = $true
        petSay          = $true
        petWeatherCity  = ''
        caiyunToken     = ''
        petOverlayW     = 0
        petOverlayH     = 0
        petFps          = 60
        petSayEvery     = 6
        petSayDur       = 0
        petOverlayFont  = 9.5
        petOverlayOpacity = 0.94
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
        if (-not $s.PSObject.Properties['pet']) { $s | Add-Member -NotePropertyName pet -NotePropertyValue $def.pet -Force }
        if (-not $s.PSObject.Properties['petX']) { $s | Add-Member -NotePropertyName petX -NotePropertyValue 0 -Force }
        if (-not $s.PSObject.Properties['petY']) { $s | Add-Member -NotePropertyName petY -NotePropertyValue 0 -Force }
        if (-not $s.PSObject.Properties['petRandomMin']) { $s | Add-Member -NotePropertyName petRandomMin -NotePropertyValue $def.petRandomMin -Force }
        if (-not $s.PSObject.Properties['petWander']) { $s | Add-Member -NotePropertyName petWander -NotePropertyValue $def.petWander -Force }
        if (-not $s.PSObject.Properties['petWeather']) { $s | Add-Member -NotePropertyName petWeather -NotePropertyValue $def.petWeather -Force }
        if (-not $s.PSObject.Properties['petScale'] -or [double]$s.petScale -le 0) { $s | Add-Member -NotePropertyName petScale -NotePropertyValue 1.0 -Force }
        if (-not $s.PSObject.Properties['petGravity']) { $s | Add-Member -NotePropertyName petGravity -NotePropertyValue 1400 -Force }
        if (-not $s.PSObject.Properties['petRestitution']) { $s | Add-Member -NotePropertyName petRestitution -NotePropertyValue 0.78 -Force }
        if (-not $s.PSObject.Properties['petGroundFriction']) { $s | Add-Member -NotePropertyName petGroundFriction -NotePropertyValue 2.5 -Force }
        if (-not $s.PSObject.Properties['petThrowPower']) { $s | Add-Member -NotePropertyName petThrowPower -NotePropertyValue 1.0 -Force }
        if (-not $s.PSObject.Properties['petCeilingBounce']) { $s | Add-Member -NotePropertyName petCeilingBounce -NotePropertyValue $true -Force }
        if (-not $s.PSObject.Properties['petSay']) { $s | Add-Member -NotePropertyName petSay -NotePropertyValue $true -Force }
        if (-not $s.PSObject.Properties['petWeatherCity']) { $s | Add-Member -NotePropertyName petWeatherCity -NotePropertyValue '' -Force }
        if (-not $s.PSObject.Properties['caiyunToken']) { $s | Add-Member -NotePropertyName caiyunToken -NotePropertyValue '' -Force }
        if (-not $s.PSObject.Properties['petOverlayW']) { $s | Add-Member -NotePropertyName petOverlayW -NotePropertyValue 0 -Force }
        if (-not $s.PSObject.Properties['petOverlayH']) { $s | Add-Member -NotePropertyName petOverlayH -NotePropertyValue 0 -Force }
        if (-not $s.PSObject.Properties['petFps']) { $s | Add-Member -NotePropertyName petFps -NotePropertyValue 60 -Force }
        if (-not $s.PSObject.Properties['petSayEvery']) { $s | Add-Member -NotePropertyName petSayEvery -NotePropertyValue 6 -Force }
        if (-not $s.PSObject.Properties['petSayDur']) { $s | Add-Member -NotePropertyName petSayDur -NotePropertyValue 0 -Force }
        if (-not $s.PSObject.Properties['petOverlayFont'] -or [double]$s.petOverlayFont -le 0) { $s | Add-Member -NotePropertyName petOverlayFont -NotePropertyValue 9.5 -Force }
        if (-not $s.PSObject.Properties['petOverlayOpacity'] -or [double]$s.petOverlayOpacity -le 0) { $s | Add-Member -NotePropertyName petOverlayOpacity -NotePropertyValue 0.94 -Force }
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
        pet             = [bool]$Settings.pet
        petX            = [int]$Settings.petX
        petY            = [int]$Settings.petY
        petRandomMin    = [int]$Settings.petRandomMin
        petWander       = [bool]$Settings.petWander
        petWeather      = [bool]$Settings.petWeather
        petScale        = [double]$(if ($Settings.petScale) { $Settings.petScale } else { 1.0 })
        petGravity      = [int]$(if ($Settings.petGravity) { $Settings.petGravity } else { 1400 })
        petRestitution  = [double]$(if ($Settings.petRestitution) { $Settings.petRestitution } else { 0.78 })
        petGroundFriction = [double]$(if ($Settings.petGroundFriction) { $Settings.petGroundFriction } else { 2.5 })
        petThrowPower   = [double]$(if ($Settings.petThrowPower) { $Settings.petThrowPower } else { 1.0 })
        petCeilingBounce = [bool]$Settings.petCeilingBounce
        petSay          = [bool]$Settings.petSay
        petWeatherCity  = [string]$Settings.petWeatherCity
        caiyunToken     = [string]$Settings.caiyunToken
        petOverlayW     = [int]$(if ($Settings.petOverlayW) { $Settings.petOverlayW } else { 0 })
        petOverlayH     = [int]$(if ($Settings.petOverlayH) { $Settings.petOverlayH } else { 0 })
        petFps          = [int]$(if ($Settings.petFps) { $Settings.petFps } else { 60 })
        petSayEvery     = [int]$(if ($null -ne $Settings.petSayEvery) { $Settings.petSayEvery } else { 6 })
        petSayDur       = [double]$(if ($null -ne $Settings.petSayDur) { $Settings.petSayDur } else { 0 })
        petOverlayFont  = [double]$(if ($Settings.petOverlayFont) { $Settings.petOverlayFont } else { 9.5 })
        petOverlayOpacity = [double]$(if ($Settings.petOverlayOpacity) { $Settings.petOverlayOpacity } else { 0.94 })
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
    # 每次只重读统计值,这样托盘每 2 秒的网速刷新就不会顶掉宠物的帧。
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
function Get-WifiInfo {
    $ssid = ''
    $state = ''
    $ifName = ''
    try { $out = & netsh wlan show interfaces 2>$null }
    catch { return [pscustomobject]@{ State = ''; Ssid = ''; Interface = '' } }

    foreach ($line in $out) {
        if ($line -match '^\s*Name\s*:\s*(.+)$' -and -not $ifName) { $ifName = $Matches[1].Trim() }
        elseif ($line -match '^\s*SSID\s*:\s*(.+)$') { $ssid = $Matches[1].Trim() }
        elseif ($line -match '^\s*(State|状态)\s*:\s*(.+)$') { $state = $Matches[2].Trim() }
    }
    return [pscustomobject]@{ State = $state; Ssid = $ssid; Interface = $ifName }
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
        foreach ($line in (& netsh wlan show profiles 2>$null)) {
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
        foreach ($line in (& netsh wlan show networks mode=bssid 2>$null)) {
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
        $r = & netsh wlan add profile (('filename="{0}"' -f $tmp)) user=current 2>&1
        $msg = (($r | Where-Object { $_ }) -join ' ').Trim()
        if ($LASTEXITCODE -eq 0) {
            Write-Log ('已自动创建开放网络配置: ' + $Ssid) 'OK'
            return $true
        }
        Write-Log ('自动创建 WiFi 配置失败(' + $LASTEXITCODE + '): ' + $msg) 'WARN'
        return $false
    } catch {
        Write-Log ('自动创建 WiFi 配置出错: ' + $_.Exception.Message) 'WARN'
        return $false
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Connect-CampusWifi {
    param([string]$Ssid = 'CIT-Chinanet', [int]$TimeoutSec = 25)

    if (Test-WifiConnected -Ssid $Ssid) {
        Write-Log ('WiFi 已连接:' + $Ssid)
        return $true
    }

    $info = Get-WifiInfo
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
    if (-not $wl.TargetInProfiles) {
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

        $r = & netsh wlan connect (('name="{0}"' -f $Ssid)) 2>&1
        $msg = (($r | Where-Object { $_ }) -join ' ').Trim()
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

# ================= 桌面宠物 =================
# 素材来自 @PC2005-cloud/dsh-pet(MIT):https://github.com/PC2005-cloud/dsh-pet
# 原版是在 Electron/WebGL 里直接放黑底 webm,一个宠物进程要 100MB 上下。
# 这里的做法:构建时把 webm 抠图(含去黑边)+ 按 24fps 逐帧导出 PNG,
# 运行时用 GDI 分层窗口(UpdateLayeredWindow)直接画在**已有的托盘进程**里。
$script:PetDir    = Join-Path $WorkDir 'pet'
$script:PetUrl    = 'https://github.com/PC2005-cloud/dsh-pet'
$script:PetCache  = @{}          # 动作 -> @{ Hbs; W; H; Count }
$script:PetForm   = $null
$script:PetTimer  = $null
$script:PetSw     = $null
$script:PetAnim   = ''
$script:PetLastIdx = -1
$script:PetAction = ''
$script:PetActionEnd = [datetime]::MinValue
$script:PetX      = 0
$script:PetY      = 0
$script:PetXf     = 0.0
$script:PetYf     = 0.0
$script:PetLastDrawX = -99999
$script:PetLastDrawY = -99999
$script:PetCfg    = $null
$script:PetCfgAt  = [datetime]::MinValue
$script:PetLastScale = 1.0
$script:PetDrag   = $null
$script:PetSettingsStamp = 0
$script:PetFpsLog  = $false
$script:PetFpsLog2 = $false
$script:PetTicks   = 0
$script:PetTicksWarm = 0
$script:PetWarmAt  = [datetime]::MinValue
$script:PetTickStart = [datetime]::MinValue
$script:PetTickAll = 0
$script:PetManifest = $null
$script:PetSleepCache = $false
$script:PetSleepAt = [datetime]::MinValue
$script:PetSleepHits = 0
$script:PetStateStamp = 0
$script:PetOverlayFontSize = 9.5
$script:PetOverlayLineH = 22
$script:PetOverlayPad = 10
$script:TrayAlive = $false
$script:TrayAliveAt = [datetime]::MinValue

# 走动 / 随机事件
$script:PetTarget = $null
$script:PetVx = 0.0
$script:PetVy = 0.0
$script:PetPhys = $false
$script:PetLastPhysAt = [datetime]::MinValue
$script:PetSayUntil = [datetime]::MinValue
$script:PetSayings = $null
$script:PetStillSince = [datetime]::MinValue
$script:PetNextRandom = [datetime]::MinValue
$script:PetMemeUntil = [datetime]::MinValue

# 悬停面板
$script:PetOverlay = $null
$script:PetOverlayHb = [IntPtr]::Zero
$script:PetOverlaySize = $null
$script:PetOverlayButtons = @()
$script:PetOverlayKind = ''
$script:PetOverlayDrawWarned = $false
$script:PetHover = $false
$script:PetHoverOffAt = [datetime]::MinValue
$script:PetOverlayHover = $false
$script:PetOverlayRefreshAt = [datetime]::MinValue

# 媒体 / 天气
$script:PetMedia = $null
$script:PetMediaAt = [datetime]::MinValue
$script:PetMediaActions = $null
$script:PetWeather = $null
$script:PetWeatherAt = [datetime]::MinValue
$script:PetWeatherFile = Join-Path $WorkDir 'pet-weather.json'
# 彩云天气接口根(测试时可用环境变量指到本地假服务)
$script:CaiyunApiBase = $(if ($env:CAMPUSNET_CAIYUN_BASE) { [string]$env:CAMPUSNET_CAIYUN_BASE } else { 'https://api.caiyunapp.com' })
$script:PetMemeCache = @{}

$script:LayeredCs = @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace CampusNetPet {
    public static class Layered {
        [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }
        [StructLayout(LayoutKind.Sequential)] public struct SIZE { public int cx; public int cy; }
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BLENDFUNCTION { public byte BlendOp; public byte BlendFlags; public byte SourceConstantAlpha; public byte AlphaFormat; }

        [DllImport("user32.dll")] static extern IntPtr GetDC(IntPtr hWnd);
        [DllImport("user32.dll")] static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("gdi32.dll")] static extern IntPtr CreateCompatibleDC(IntPtr hDC);
        [DllImport("gdi32.dll")] static extern bool DeleteDC(IntPtr hDC);
        [DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);
        [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr hObject);
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr value);
        [DllImport("user32.dll")]
        static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref POINT pptDst, ref SIZE psize,
            IntPtr hdcSrc, ref POINT pptSrc, int crKey, ref BLENDFUNCTION pblend, int dwFlags);

        public static bool MakeDpiAware() {
            try { return SetProcessDpiAwarenessContext(new IntPtr(-4)); }   // PER_MONITOR_AWARE_V2
            catch { return false; }
        }

        // WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
        public static void Enable(IntPtr hwnd) {
            int ex = GetWindowLong(hwnd, -20);
            SetWindowLong(hwnd, -20, ex | 0x80000 | 0x80 | 0x8000000);
        }

        [DllImport("winmm.dll")] static extern uint timeBeginPeriod(uint ms);
        [DllImport("winmm.dll")] static extern uint timeEndPeriod(uint ms);
        public static void BeginHighRes() { try { timeBeginPeriod(1); } catch { } }
        public static void EndHighRes() { try { timeEndPeriod(1); } catch { } }

        [StructLayout(LayoutKind.Sequential)]
        struct POWER_THROTTLING_STATE { public uint Version; public uint ControlMask; public uint StateMask; }
        [DllImport("kernel32.dll")] static extern IntPtr GetCurrentProcess();
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetProcessInformation(IntPtr hProcess, int infoClass, ref POWER_THROTTLING_STATE state, uint size);

        public static bool DisablePowerThrottling() {
            try {
                var st = new POWER_THROTTLING_STATE();
                st.Version = 1; st.ControlMask = 0x1; st.StateMask = 0;
                return SetProcessInformation(GetCurrentProcess(), 4, ref st, (uint)Marshal.SizeOf(st));
            } catch { return false; }
        }

        public static IntPtr ToHBitmap(Bitmap bmp) { return bmp.GetHbitmap(Color.FromArgb(0)); }
        public static void FreeHBitmap(IntPtr h) { if (h != IntPtr.Zero) DeleteObject(h); }

        [ThreadStatic] public static int LastError;

        public static bool Draw(IntPtr hwnd, IntPtr hbmp, int w, int h, int x, int y, byte alpha) {
            IntPtr screenDc = GetDC(IntPtr.Zero);
            IntPtr memDc = CreateCompatibleDC(screenDc);
            IntPtr old = SelectObject(memDc, hbmp);
            try {
                SIZE size = new SIZE(); size.cx = w; size.cy = h;
                POINT src = new POINT();
                POINT dst = new POINT(); dst.X = x; dst.Y = y;
                BLENDFUNCTION blend = new BLENDFUNCTION();
                blend.BlendOp = 0; blend.BlendFlags = 0;
                blend.SourceConstantAlpha = alpha; blend.AlphaFormat = 1;
                bool ok = UpdateLayeredWindow(hwnd, screenDc, ref dst, ref size, memDc, ref src, 0, ref blend, 2);
                if (!ok) LastError = Marshal.GetLastWin32Error();
                return ok;
            } finally {
                SelectObject(memDc, old);
                DeleteDC(memDc);
                ReleaseDC(IntPtr.Zero, screenDc);
            }
        }
    }
}
'@

function Test-PetAssets {
    if (-not (Test-Path (Join-Path $script:PetDir 'idle'))) { return $false }
    return $true
}

function Set-PetDpiAware {
    if (-not ('CampusNetPet.Dpi' -as [type])) {
        try {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace CampusNetPet {
    public static class Dpi {
        [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr v);
        public static bool Make() {
            try { return SetProcessDpiAwarenessContext(new IntPtr(-4)); } catch { return false; }
        }
    }
}
'@
        } catch { }
    }
    try { return [CampusNetPet.Dpi]::Make() } catch { return $false }
}

function Get-PetManifest {
    if ($script:PetManifest) { return $script:PetManifest }
    $f = Join-Path $script:PetDir 'pet.json'
    if (-not (Test-Path $f)) { return $null }
    try {
        $script:PetManifest = [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        return $script:PetManifest
    } catch { return $null }
}

function Get-PetFps { $m = Get-PetManifest; if ($m -and [int]$m.fps -ge 1) { return [int]$m.fps }; return 24 }
function Get-PetAnimList { $m = Get-PetManifest; if (-not $m -or -not $m.anims) { return @() }; return @($m.anims) }
function Get-PetMemeList { $m = Get-PetManifest; if (-not $m -or -not $m.memes) { return @() }; return @($m.memes) }

function Get-PetLabel {
    param([string]$Id)
    foreach ($a in (Get-PetAnimList)) { if ($a.id -eq $Id) { return [string]$a.label } }
    return $Id
}

function Merge-PetBitmap {
    # 两张同尺寸帧线性混合(0.5 就是中间帧),用来把 24fps 素材插成 48fps
    param($A, $B, [double]$Alpha = 0.5)
    if (-not $A -or -not $B) { return $null }
    if ($A.Width -ne $B.Width -or $A.Height -ne $B.Height) { return $null }
    $w = $A.Width; $h = $A.Height
    $out = $null
    try {
        $out = New-Object System.Drawing.Bitmap -ArgumentList $w, $h, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g = [System.Drawing.Graphics]::FromImage($out)
        $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.Clear([System.Drawing.Color]::Transparent)
        $g.DrawImage($A, 0, 0, $w, $h)

        $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
        $ia = New-Object System.Drawing.Imaging.ImageAttributes
        $cm = New-Object System.Drawing.Imaging.ColorMatrix
        $cm.Matrix33 = [single]$Alpha
        $ia.SetColorMatrix($cm)
        $rect = New-Object System.Drawing.Rectangle -ArgumentList 0, 0, $w, $h
        $g.DrawImage($B, $rect, 0, 0, $w, $h, [System.Drawing.GraphicsUnit]::Pixel, $ia)
        $ia.Dispose()
        $g.Dispose()
        return $out
    } catch {
        if ($out) { try { $out.Dispose() } catch { } }
        return $null
    }
}

function Get-PetFrames {
    # 按缩放比例缓存:key = "动作@比例"。换大小只会重新解码当前动作,不影响别的
    param([string]$Anim, [double]$Scale = 1.0)

    if ($Scale -lt 0.5) { $Scale = 0.5 }
    if ($Scale -gt 2.0) { $Scale = 2.0 }
    $key = ('{0}@{1:N2}' -f $Anim, $Scale)
    if ($script:PetCache.ContainsKey($key)) { return $script:PetCache[$key] }

    $dir = Join-Path $script:PetDir $Anim
    if (-not (Test-Path $dir)) { return $null }
    $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.png' | Sort-Object Name)
    if ($files.Count -eq 0) { return $null }

    $fps = Get-PetFps
    foreach ($a in (Get-PetAnimList)) {
        if ($a.id -eq $Anim -and $a.PSObject.Properties['fps'] -and [int]$a.fps -gt 0) { $fps = [int]$a.fps }
    }

    $hbs = New-Object System.Collections.ArrayList
    $w = 0; $h = 0

    # 素材是 24fps;想要更高刷新率就插中间帧(线性混合),播放时长不变
    $targetFps = [int](Get-PetCfg).Fps
    if ($targetFps -lt 12) { $targetFps = 12 }
    $k = 1
    if ($targetFps -gt [int]($fps * 1.2)) { $k = 2 }
    # 内存闸门:一个动作最多 300 帧(约 80MB),再高就不插了
    if (($files.Count * $k) -gt 300) {
        Write-Log ('帧插值跳过:' + $Anim + ' 帧数偏多(' + $files.Count + ' × ' + $k + ')') 'WARN'
        $k = 1
    }

    $prev = $null
    foreach ($f in $files) {
        $bmp = $null
        try {
            $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
            $ms = New-Object System.IO.MemoryStream -ArgumentList (,$bytes)
            $img = [System.Drawing.Image]::FromStream($ms)
            $raw = New-Object System.Drawing.Bitmap -ArgumentList $img
            $img.Dispose(); $ms.Dispose()

            if ([Math]::Abs($Scale - 1.0) -lt 0.01) {
                $bmp = $raw
            } else {
                $dw = [Math]::Max(1, [int][Math]::Round($raw.Width * $Scale))
                $dh = [Math]::Max(1, [int][Math]::Round($raw.Height * $Scale))
                $scaled = New-Object System.Drawing.Bitmap -ArgumentList $dw, $dh, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $g = [System.Drawing.Graphics]::FromImage($scaled)
                $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
                $g.Clear([System.Drawing.Color]::Transparent)
                $g.DrawImage($raw, 0, 0, $dw, $dh)
                $g.Dispose()
                $raw.Dispose()
                $bmp = $scaled
            }
            if ($w -eq 0) { $w = $bmp.Width; $h = $bmp.Height }

            # 上一帧和这一帧之间插一张(顺序:帧0 → 混合 → 帧1 → …)
            if ($k -gt 1 -and $prev -and $prev.Width -eq $bmp.Width -and $prev.Height -eq $bmp.Height) {
                $mid = Merge-PetBitmap -A $prev -B $bmp -Alpha 0.5
                if ($mid) {
                    [void]$hbs.Add([CampusNetPet.Layered]::ToHBitmap($mid))
                    $mid.Dispose()
                }
            }
            [void]$hbs.Add([CampusNetPet.Layered]::ToHBitmap($bmp))

            if ($prev) { $prev.Dispose() }
            $prev = $bmp
        } catch {
            if ($bmp) { try { $bmp.Dispose() } catch { } }
        }
    }
    if ($prev) { try { $prev.Dispose() } catch { } }
    if ($hbs.Count -eq 0) { return $null }

    $entry = [pscustomobject]@{ Hbs = $hbs; W = $w; H = $h; Count = $hbs.Count; Fps = ($fps * $k) }
    $script:PetCache[$key] = $entry
    if ($script:PetCache.Count -gt 3) {
        foreach ($k in @($script:PetCache.Keys)) {
            if ($k -eq $key -or $k -like 'idle@*' -or $k -like 'walk@*') { continue }
            foreach ($hb in $script:PetCache[$k].Hbs) { [CampusNetPet.Layered]::FreeHBitmap($hb) }
            $script:PetCache.Remove($k)
        }
    }
    return $entry
}

function Test-PetShouldSleep {
    # 每秒读一次 state.json(要解析 JSON,好几毫秒)会顶掉一帧 —— 掉帧的另一半原因。
    # 改成 3 秒一次,而且文件没变就不解析。
    if (((Get-Date) - $script:PetSleepAt).TotalSeconds -lt 3) { return $script:PetSleepCache }
    $script:PetSleepAt = Get-Date

    $stamp = 0
    try { if (Test-Path $StateFn) { $stamp = (Get-Item $StateFn).LastWriteTimeUtc.Ticks } } catch { }
    if ($stamp -eq $script:PetStateStamp) { return $script:PetSleepCache }
    $script:PetStateStamp = $stamp

    $st = Read-State
    if (-not $st) { return $script:PetSleepCache }
    $sleepy = $false
    if (-not $st.online) { $sleepy = $true }
    elseif ((Get-StateAge -State $st) -ge ($IntervalSec * 4)) { $sleepy = $true }
    if ($sleepy) { $script:PetSleepHits++ } else { $script:PetSleepHits = 0 }
    $script:PetSleepCache = ($script:PetSleepHits -ge 2)
    return $script:PetSleepCache
}

function Get-PetCfg {
    # 每帧都要用的几个设置:缓存 2 秒。之前每帧都 Get-Settings(读磁盘 + 解析 JSON),
    # 60 次/秒 —— 这是掉帧和移动发涩的主要来源。
    if ($script:PetCfg -and ((Get-Date) - $script:PetCfgAt).TotalSeconds -lt 2) { return $script:PetCfg }
    $s = $null
    try { $s = Get-Settings } catch { }
    $scale = 1.0
    if ($s -and $s.PSObject.Properties['petScale']) { $scale = [double]$s.petScale }
    if ($scale -lt 0.5) { $scale = 0.5 }
    if ($scale -gt 2.0) { $scale = 2.0 }
    $script:PetCfg = [pscustomobject]@{
        RandomMin = $(if ($s) { [int]$s.petRandomMin } else { 5 })
        Wander    = $(if ($s) { [bool]$s.petWander } else { $true })
        Weather   = $(if ($s) { [bool]$s.petWeather } else { $true })
        Scale     = $scale
        OverlayFont    = $(if ($s -and $s.PSObject.Properties['petOverlayFont'] -and [double]$s.petOverlayFont -gt 0) { [double]$s.petOverlayFont } else { 9.5 })
        OverlayOpacity = $(if ($s -and $s.PSObject.Properties['petOverlayOpacity'] -and [double]$s.petOverlayOpacity -gt 0) { [double]$s.petOverlayOpacity } else { 0.94 })
        OverlayW       = $(if ($s -and $s.PSObject.Properties['petOverlayW'] -and [int]$s.petOverlayW -gt 0) { [int]$s.petOverlayW } else { 0 })
        Fps            = $(if ($s -and $s.PSObject.Properties['petFps'] -and [int]$s.petFps -ge 12) { [int]$s.petFps } else { 60 })
        SayEvery       = $(if ($s -and $s.PSObject.Properties['petSayEvery']) { [int]$s.petSayEvery } else { 6 })
        SayDur         = $(if ($s -and $s.PSObject.Properties['petSayDur']) { [double]$s.petSayDur } else { 0 })
        OverlayH       = $(if ($s -and $s.PSObject.Properties['petOverlayH'] -and [int]$s.petOverlayH -gt 0) { [int]$s.petOverlayH } else { 0 })
        Gravity        = $(if ($s -and $s.PSObject.Properties['petGravity'] -and [int]$s.petGravity -gt 0) { [double]$s.petGravity } else { 1400 })
        Restitution    = $(if ($s -and $s.PSObject.Properties['petRestitution']) { [double]$s.petRestitution } else { 0.78 })
        Friction       = $(if ($s -and $s.PSObject.Properties['petGroundFriction']) { [double]$s.petGroundFriction } else { 2.5 })
        ThrowPower     = $(if ($s -and $s.PSObject.Properties['petThrowPower'] -and [double]$s.petThrowPower -gt 0) { [double]$s.petThrowPower } else { 1.0 })
        CeilingBounce  = $(if ($s -and $s.PSObject.Properties['petCeilingBounce']) { [bool]$s.petCeilingBounce } else { $true })
        Say            = $(if ($s -and $s.PSObject.Properties['petSay']) { [bool]$s.petSay } else { $true })
        WeatherCity    = $(if ($s -and $s.PSObject.Properties['petWeatherCity']) { [string]$s.petWeatherCity } else { '' })
    }
    $script:PetCfgAt = Get-Date
    return $script:PetCfg
}

function Set-PetSize {
    param([double]$Scale)
    if ($Scale -lt 0.5) { $Scale = 0.5 }
    if ($Scale -gt 2.0) { $Scale = 2.0 }
    try {
        $s = Get-Settings
        $s | Add-Member -NotePropertyName petScale -NotePropertyValue $Scale -Force
        Save-Settings -Settings $s
    } catch { }
    $script:PetCfg = $null
    $script:PetCfgAt = [datetime]::MinValue
    Write-Log ('宠物大小调整为 ' + [int][Math]::Round($Scale * 100) + '%')
}

function Set-PetOverlaySize {
    # 信息框尺寸:0 = 自动(按内容);宽固定就换行,高固定就截断
    param([int]$W, [int]$H)
    if ($W -lt 0) { $W = 0 }
    if ($H -lt 0) { $H = 0 }
    try {
        $s = Get-Settings
        $s.petOverlayW = $W
        $s.petOverlayH = $H
        Save-Settings -Settings $s
    } catch { }
    $script:PetCfg = $null
    $script:PetCfgAt = [datetime]::MinValue
    Write-Log ('信息框尺寸: ' + $(if ($W -gt 0) { [string]$W } else { '自动' }) + ' × ' + $(if ($H -gt 0) { [string]$H } else { '自动' }))
    # 面板正开着就按新尺寸重画一次
    if ($script:PetOverlayKind -eq 'info' -and $script:PetOverlay -and $script:PetOverlay.Visible) {
        Set-PetOverlayContent -Kind 'info'
    }
}

function Get-PetSetting {
    param([string]$Name, $Default)
    try {
        $s = Get-Settings
        $p = $s.PSObject.Properties[$Name]
        if ($p -and $null -ne $p.Value) { return $p.Value }
    } catch { }
    return $Default
}

function Set-PetSettingValue {
    param([string]$Name, $Value)
    try {
        $s = Get-Settings
        $s | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
        Save-Settings -Settings $s
    } catch { }
}

# ---------------- 天气(联网识别位置) ----------------
function Get-WeatherText {
    param([int]$Code)
    if ($Code -eq 0) { return '晴' }
    if ($Code -in 1, 2) { return '多云' }
    if ($Code -eq 3) { return '阴' }
    if ($Code -in 45, 48) { return '雾' }
    if ($Code -ge 51 -and $Code -le 57) { return '毛毛雨' }
    if ($Code -ge 61 -and $Code -le 67) { return '雨' }
    if ($Code -ge 71 -and $Code -le 77) { return '雪' }
    if ($Code -ge 80 -and $Code -le 82) { return '阵雨' }
    if ($Code -ge 85 -and $Code -le 86) { return '阵雪' }
    if ($Code -ge 95) { return '雷雨' }
    return '—'
}

function Get-CaiyunSkyText {
    # 彩云天气的 skycon 枚举 → 中文
    param([string]$Skycon)
    switch ([string]$Skycon) {
        'CLEAR_DAY' { return '晴' }
        'CLEAR_NIGHT' { return '晴' }
        'PARTLY_CLOUDY_DAY' { return '多云' }
        'PARTLY_CLOUDY_NIGHT' { return '多云' }
        'CLOUDY' { return '阴' }
        'LIGHT_HAZE' { return '轻度霾' }
        'MODERATE_HAZE' { return '中度霾' }
        'HEAVY_HAZE' { return '重度霾' }
        'LIGHT_RAIN' { return '小雨' }
        'MODERATE_RAIN' { return '中雨' }
        'HEAVY_RAIN' { return '大雨' }
        'STORM_RAIN' { return '暴雨' }
        'FOG' { return '雾' }
        'LIGHT_SNOW' { return '小雪' }
        'MODERATE_SNOW' { return '中雪' }
        'HEAVY_SNOW' { return '大雪' }
        'STORM_SNOW' { return '暴雪' }
        'DUST' { return '浮尘' }
        'SAND' { return '扬沙' }
        'WIND' { return '大风' }
        default { return '—' }
    }
}

function Get-DressingAdvice {
    # 穿衣提醒:按体感温度分档,再按天气加一句提醒
    param([double]$Temp, [string]$Text = '', [double]$WindSpeed = 0)

    $t = [int][Math]::Round($Temp)
    $adv = ''
    if ($t -ge 30) { $adv = '短袖短裤就行' }
    elseif ($t -ge 26) { $adv = '短袖 T 恤' }
    elseif ($t -ge 22) { $adv = '薄长袖,或短袖加件薄外套' }
    elseif ($t -ge 18) { $adv = '长袖衬衫 / 薄卫衣' }
    elseif ($t -ge 14) { $adv = '卫衣 / 夹克' }
    elseif ($t -ge 10) { $adv = '外套加长裤' }
    elseif ($t -ge 5) { $adv = '厚外套 / 毛衣' }
    elseif ($t -ge 0) { $adv = '棉服 / 羽绒服' }
    else { $adv = '羽绒服,帽子围巾戴上' }

    $extra = @()
    if ($Text -match '雨') { $extra += '带伞' }
    if ($Text -match '雪') { $extra += '带伞,路滑' }
    if ($Text -match '霾|浮尘|扬沙') { $extra += '戴口罩' }
    if ($WindSpeed -ge 5.5) { $extra += '风大' }
    if ($t -ge 30) { $extra += '多喝水' }

    if ($extra.Count -gt 0) { return ($adv + ',' + ($extra -join ',')) }
    return $adv
}

function Get-CnPlaceCandidates {
    # 「吉林省长春市南关区」这种整串,地理编码基本都不认;拆成几级逐个试
    param([string]$Text)
    $t = ([string]$Text).Trim()
    $list = @()
    if (-not $t) { return $list }
    $list += $t
    if ($t -match '([^省]+省)?(.+)$') { }
    # 去掉末尾的 省/市/区/县
    $a = $t -replace '(特别行政区|自治区|省|市|区|县)$', ''
    if ($a -and $a -ne $t) { $list += $a }
    # 「省」后面的部分(市 + 区县)
    if ($t -match '省(.+)$') {
        $b = $Matches[1]
        if ($b) { $list += $b }
        $b2 = $b -replace '(市|区|县)$', ''
        if ($b2 -and $b2 -ne $b) { $list += $b2 }
    }
    # 「市」后面的部分(区县)
    if ($t -match '市(.+)$') {
        $c = $Matches[1]
        if ($c) { $list += $c }
    }
    return @($list | Where-Object { $_ } | Select-Object -Unique)
}

function Get-CnPlaceParts {
    # 从 Nominatim 的 address 字段 / display_name 里按后缀归类成 省 / 市 / 区县
    # (display_name 的层级顺序不固定,不能直接按顺序取)
    param([string[]]$Texts)
    $prov = ''; $city = ''; $dist = ''
    foreach ($t in $Texts) {
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        foreach ($seg in ([string]$t -split ',')) {
            $s = $seg.Trim()
            if (-not $s) { continue }
            if ($s -match '^(中国|China|中华人民共和国|People.s Republic of China)$') { continue }
            if ($s -match '^\d+$') { continue }
            if ((-not $prov) -and $s -match '(省|自治区|特别行政区)$') { $prov = $s; continue }
            if ((-not $city) -and $s -match '(市|自治州|地区|盟)$') { $city = $s; continue }
            if ((-not $dist) -and $s -match '(区|县|旗)$') { $dist = $s; continue }
        }
    }
    return [pscustomobject]@{ Province = $prov; City = $city; District = $dist }
}

function Add-CnAdminSuffix {
    param([string]$Name, [string]$Kind)
    $n = ([string]$Name).Trim()
    if (-not $n) { return '' }
    if ($Kind -eq 'province') {
        if ($n -match '(省|自治区|特别行政区)$') { return $n }
        if (@('北京', '上海', '天津', '重庆') -contains $n) { return $n }
        return ($n + '省')
    }
    if ($Kind -eq 'city') {
        if ($n -match '(市|州|盟|地区|自治州|自治县)$') { return $n }
        return ($n + '市')
    }
    return $n
}

function Get-WeatherPlace {
    # 省 / 市 / 区县 —— 既用来显示,也给天气接口经纬度
    # 顺序:Nominatim(能到区县)→ Open-Meteo(只到市)→ IP 定位(只到市)
    param([string]$WantCity = '')

    $r = [pscustomobject]@{ Label = ''; Province = ''; City = ''; District = ''; Lat = $null; Lon = $null; Ok = $false }

    if ($WantCity) {
        # 1) Nominatim:地址层级全,能拿到区县
        foreach ($cand in (Get-CnPlaceCandidates $WantCity)) {
            if ($r.Ok) { break }
            try {
                $u = 'https://nominatim.openstreetmap.org/search?q={0}&format=json&addressdetails=1&limit=1&accept-language=zh-CN' -f [uri]::EscapeDataString($cand)
                $nr = Invoke-RestMethod -Uri $u -TimeoutSec 12 -UseBasicParsing -Headers @{ 'User-Agent' = ('campus-net/' + $script:AppVersion + ' (personal use)') }
                if ($nr -and @($nr).Count -gt 0) {
                    $one = @($nr)[0]
                    $r.Lat = [double]$one.lat
                    $r.Lon = [double]$one.lon
                    $ad = $one.address
                    $segs = @()
                    if ($ad) {
                        foreach ($fld in @($ad.state, $ad.city, $ad.town, $ad.municipality, $ad.district, $ad.county, $ad.city_district, $ad.suburb, $ad.village)) {
                            if ($fld) { $segs += [string]$fld }
                        }
                    }
                    $segs += [string]$one.display_name
                    $pp = Get-CnPlaceParts -Texts $segs
                    $r.Province = Add-CnAdminSuffix ([string]$pp.Province) 'province'
                    $r.City = [string]$pp.City
                    $r.District = [string]$pp.District
                    $lbl = @()
                    foreach ($x in @($r.Province, $r.City, $r.District)) {
                        $x = ([string]$x).Trim()
                        if ($x -and ($lbl -notcontains $x)) { $lbl += $x }
                    }
                    if ($lbl.Count -gt 0) { $r.Label = ($lbl -join ' ') }
                    $r.Ok = $true
                }
            } catch {
                Write-Log ('Nominatim 定位失败(' + $cand + '):' + $_.Exception.Message) 'WARN'
            }
        }

        # 2) Open-Meteo 兜底(没有区县)
        if (-not $r.Ok) {
            foreach ($cand in (Get-CnPlaceCandidates $WantCity)) {
                if ($r.Ok) { break }
                try {
                    $u = 'https://geocoding-api.open-meteo.com/v1/search?name={0}&count=1&language=zh&format=json' -f [uri]::EscapeDataString($cand)
                    $g = Invoke-RestMethod -Uri $u -TimeoutSec 12 -UseBasicParsing
                    if ($g.results -and @($g.results).Count -gt 0) {
                        $one = @($g.results)[0]
                        $r.Lat = [double]$one.latitude
                        $r.Lon = [double]$one.longitude
                        $r.Province = Add-CnAdminSuffix ([string]$one.admin1) 'province'
                        $r.City = Add-CnAdminSuffix ([string]$one.admin2) 'city'
                        if (-not $one.admin2) { $r.City = [string]$one.name }
                        $r.District = [string]$one.admin3
                        $r.Ok = $true
                    }
                } catch {
                    Write-Log ('天气地理编码失败(' + $cand + '):' + $_.Exception.Message) 'WARN'
                }
            }
        }
    }

    if (-not $r.Ok) {
        # 3) IP 定位只能到市,没有区县
        try {
            $loc = Invoke-RestMethod -Uri 'http://ip-api.com/json/?fields=status,country,regionName,city,lat,lon' -TimeoutSec 12 -UseBasicParsing
            if ($loc.status -eq 'success') {
                $r.Lat = [double]$loc.lat
                $r.Lon = [double]$loc.lon
                $r.Province = [string]$loc.regionName
                $r.City = [string]$loc.city
                $r.Ok = $true
            }
        } catch {
            Write-Log ('IP 定位失败:' + $_.Exception.Message) 'WARN'
        }
    }

    # 没有现成标签就按 省/市/区 拼(同名只留一个)
    if (-not $r.Label) {
        $parts = @()
        foreach ($p in @($r.Province, $r.City, $r.District)) {
            $p = ([string]$p).Trim()
            if (-not $p) { continue }
            $dup = $false
            foreach ($q in $parts) {
                if ($q -eq $p -or $q -like ('*' + $p + '*') -or $p -like ('*' + $q + '*')) { $dup = $true; break }
            }
            if (-not $dup) { $parts += $p }
        }
        $r.Label = ($parts -join ' ')
    }
    return $r
}

function Get-CaiyunWeather {
    # 彩云天气 v2.6:一次 weather.json 拿实时 + 穿衣指数 + 预警
    param([string]$Token, [double]$Lat, [double]$Lon)

    $r = [pscustomobject]@{ Ok = $false; Temp = $null; Text = ''; Feels = $null; Wind = 0; Dressing = ''; Alert = '' }
    if ([string]::IsNullOrWhiteSpace($Token) -or $null -eq $Lat) { return $r }

    $coord = ('{0},{1}' -f [Math]::Round($Lat, 4), [Math]::Round($Lon, 4))
    $u = '{0}/v2.6/{1}/{2}/weather.json?alert=true&dailysteps=1' -f $script:CaiyunApiBase, $Token, $coord
    try { $w = Invoke-RestMethod -Uri $u -TimeoutSec 15 -UseBasicParsing } catch { return $r }
    if ([string]$w.status -ne 'ok') { return $r }

    $rt = $w.result.realtime
    if (-not $rt) { return $r }

    $r.Temp = [double]$rt.temperature
    $r.Text = Get-CaiyunSkyText -Skycon ([string]$rt.skycon)
    if ($null -ne $rt.apparent_temperature) { $r.Feels = [double]$rt.apparent_temperature }
    if ($rt.wind -and $null -ne $rt.wind.speed) { $r.Wind = [double]$rt.wind.speed }

    # 彩云自带的穿衣指数(desc 是"较冷 / 舒适"这种等级),拿来做补充
    try {
        $d = $w.result.daily.life_index.dressing
        if ($d) {
            if (@($d).Count -gt 0) { $r.Dressing = [string](@($d)[0].desc) }
            else { $r.Dressing = [string]$d.desc }
        }
    } catch { }

    try {
        $al = $w.result.alert.content
        if ($al -and @($al).Count -gt 0) {
            $first = @($al)[0]
            if ($first -and $first.title) { $r.Alert = [string]$first.title }
        }
    } catch { }

    $r.Ok = $true
    return $r
}

function Update-PetWeather {
    # 位置:设置里填的城市(能拿到省/市/区)→ 没填就 IP 定位(只有市)
    # 天气:彩云天气(要 token)→ 失败或没 token 用 Open-Meteo 兜底
    $out = [ordered]@{
        ts = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); ok = $false
        city = ''; place = ''; province = ''; district = ''
        temp = $null; feels = $null; code = -1; text = ''
        dressing = ''; wind = 0; alert = ''; source = ''
    }

    $wantCity = ''
    $token = ''
    try {
        $st = Get-Settings
        $wantCity = [string]$st.petWeatherCity
        $token = [string]$st.caiyunToken
    } catch { }
    $wantCity = $wantCity.Trim()
    $token = $token.Trim()

    $p = Get-WeatherPlace -WantCity $wantCity
    $out.city = [string]$p.City
    $out.place = [string]$p.Label
    $out.province = [string]$p.Province
    $out.district = [string]$p.District

    # 1) 彩云天气
    if ($token -and $p.Ok) {
        $cy = Get-CaiyunWeather -Token $token -Lat $p.Lat -Lon $p.Lon
        if ($cy.Ok) {
            $out.temp = [int][Math]::Round([double]$cy.Temp)
            $out.feels = $(if ($null -ne $cy.Feels) { [int][Math]::Round([double]$cy.Feels) } else { $out.temp })
            $out.text = [string]$cy.Text
            $out.wind = [double]$cy.Wind
            $out.alert = [string]$cy.Alert
            $base = Get-DressingAdvice -Temp $out.feels -Text $out.text -WindSpeed $out.wind
            $out.dressing = $(if ($cy.Dressing) { $base + '(' + $cy.Dressing + ')' } else { $base })
            $out.source = '彩云天气'
            $out.ok = $true
        } else {
            Write-Log '彩云天气取数失败(检查 token / 城市),回退 Open-Meteo' 'WARN'
        }
    }

    # 2) Open-Meteo 兜底
    if (-not $out.ok -and $p.Ok) {
        try {
            $u = 'https://api.open-meteo.com/v1/forecast?latitude={0}&longitude={1}&current=temperature_2m,apparent_temperature,weather_code,wind_speed_10m' -f $p.Lat, $p.Lon
            $w = Invoke-RestMethod -Uri $u -TimeoutSec 12 -UseBasicParsing
            $out.temp = [int][math]::Round([double]$w.current.temperature_2m)
            $out.code = [int]$w.current.weather_code
            $out.text = Get-WeatherText -Code $out.code
            if ($null -ne $w.current.apparent_temperature) { $out.feels = [int][Math]::Round([double]$w.current.apparent_temperature) }
            if ($null -ne $w.current.wind_speed_10m) { $out.wind = [double]$w.current.wind_speed_10m / 3.6 }
            if ($null -eq $out.feels) { $out.feels = $out.temp }
            $out.dressing = Get-DressingAdvice -Temp $out.feels -Text $out.text -WindSpeed $out.wind
            $out.source = 'Open-Meteo'
            $out.ok = $true
        } catch {
            Write-Log ('Open-Meteo 取数失败:' + $_.Exception.Message) 'WARN'
        }
    }

    try { ($out | ConvertTo-Json -Compress) | Set-Content -Path $script:PetWeatherFile -Encoding UTF8 } catch { }
    return $out
}

function Get-PetWeather {
    # 30 分钟内复用缓存
    if ($script:PetWeather -and ((Get-Date) - $script:PetWeatherAt).TotalMinutes -lt 30) { return $script:PetWeather }
    if (-not (Test-Path $script:PetWeatherFile)) { return $null }
    try {
        $w = Get-Content $script:PetWeatherFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $script:PetWeather = $w
        $script:PetWeatherAt = Get-Date
        if ($w.ts) {
            try {
                $age = ((Get-Date) - [datetime]::ParseExact($w.ts, 'yyyy-MM-dd HH:mm:ss', $null)).TotalMinutes
                if ($age -gt 30) { $script:PetWeatherAt = (Get-Date).AddMinutes(-31) }
            } catch { }
        }
        return $w
    } catch { return $null }
}

function Format-PetWeather {
    $w = Get-PetWeather
    if (-not $w -or -not $w.ok) { return '' }
    $line = [string]$w.text + ' ' + [string]$w.temp + '°C'
    if ($w.place) { $line += ' · ' + [string]$w.place }
    elseif ($w.city) { $line += ' · ' + [string]$w.city }
    if ($w.dressing) { $line += ' · ' + [string]$w.dressing }
    return $line
}

# ---------------- 媒体(GSMTC) ----------------
function Initialize-WinRtAsync {
    if ($script:PetMediaActions) { return $true }
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
        $methods = [System.WindowsRuntimeSystemExtensions].GetMethods()
        $script:PetMediaActions = @($methods | Where-Object {
                $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
            })[0]
        return [bool]$script:PetMediaActions
    } catch { return $false }
}

function Await-WinRt {
    param($Op, [Type]$ResultType, [int]$TimeoutMs = 1500)
    if (-not (Initialize-WinRtAsync)) { return $null }
    try {
        $m = $script:PetMediaActions.MakeGenericMethod($ResultType)
        $task = $m.Invoke($null, @($Op))
        if (-not $task.Wait($TimeoutMs)) { return $null }
        return $task.Result
    } catch { return $null }
}

function Get-PetMedia {
    # 返回 @{ Ok; Title; Artist; Playing; App } ;没有播放器时 Ok=$false
    if (((Get-Date) - $script:PetMediaAt).TotalSeconds -lt 1) { return $script:PetMedia }
    $script:PetMediaAt = Get-Date

    $res = [pscustomobject]@{ Ok = $false; Title = ''; Artist = ''; Playing = $false; App = '' }
    try {
        $null = [Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager, Windows.Media.Control, ContentType = WindowsRuntime]
        $op = [Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager]::RequestAsync()
        $mgr = Await-WinRt -Op $op -ResultType ([Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager])
        if ($mgr) {
            $sess = $mgr.GetCurrentSession()
            if ($sess) {
                $props = Await-WinRt -Op $sess.TryGetMediaPropertiesAsync() -ResultType ([Windows.Media.Control.GlobalSystemMediaTransportControlsSessionMediaProperties])
                if ($props) {
                    $res.Ok = $true
                    $res.Title = [string]$props.Title
                    $res.Artist = [string]$props.Artist
                    $res.App = [string]$props.SourceAppUserModelId
                    $info = $sess.GetPlaybackInfo()
                    $res.Playing = ([int]$info.PlaybackStatus -eq 4)   # Playing
                }
            }
        }
    } catch { }
    $script:PetMedia = $res
    return $res
}

function Send-PetMediaCommand {
    param([ValidateSet('playpause', 'next', 'prev')][string]$Action)
    try {
        $null = [Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager, Windows.Media.Control, ContentType = WindowsRuntime]
        $op = [Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager]::RequestAsync()
        $mgr = Await-WinRt -Op $op -ResultType ([Windows.Media.Control.GlobalSystemMediaTransportControlsSessionManager])
        if (-not $mgr) { return $false }
        $sess = $mgr.GetCurrentSession()
        if (-not $sess) { return $false }
        switch ($Action) {
            'playpause' { $null = $sess.TryTogglePlayPauseAsync() }
            'next' { $null = $sess.TrySkipNextAsync() }
            'prev' { $null = $sess.TrySkipPreviousAsync() }
        }
        $script:PetMediaAt = [datetime]::MinValue
        return $true
    } catch { return $false }
}

function Format-PetMediaLine {
    $m = Get-PetMedia
    if (-not $m -or -not $m.Ok) { return '' }
    $t = $m.Title
    if ($m.Artist) { $t = $t + ' - ' + $m.Artist }
    return $t
}

# ---------------- 悬停面板 / 表情包 ----------------
function Get-PetMemeImage {
    param([string]$Id)
    if ($script:PetMemeCache.ContainsKey($Id)) { return $script:PetMemeCache[$Id] }
    $f = Join-Path (Join-Path $script:PetDir 'memes') ($Id + '.png')
    if (-not (Test-Path $f)) { return $null }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($f)
        $ms = New-Object System.IO.MemoryStream -ArgumentList (,$bytes)
        $img = [System.Drawing.Image]::FromStream($ms)
        $bmp = New-Object System.Drawing.Bitmap -ArgumentList $img
        $img.Dispose(); $ms.Dispose()
        $script:PetMemeCache[$Id] = $bmp
        return $bmp
    } catch { return $null }
}

function Show-PetMeme {
    param([string]$Id)
    $bmp = Get-PetMemeImage -Id $Id
    if (-not $bmp) { return }
    $script:PetMemeUntil = (Get-Date).AddSeconds(8)
    Set-PetOverlayContent -Kind 'meme' -Meme $bmp
}

function New-PetOverlayWindow {
    if ($script:PetOverlay -and -not $script:PetOverlay.IsDisposed) { return }
    $f = New-Object System.Windows.Forms.Form
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $f.ShowInTaskbar = $false
    $f.TopMost = $true
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $f.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $f.Text = '校园网宠物信息'
    $f.ClientSize = New-Object System.Drawing.Size -ArgumentList 10, 10
    # 用普通窗口自绘 + 整体不透明度,而不是 UpdateLayeredWindow:
    # 半透明面板用 ULW 在跨窗口场景下容易失败(实测返回 0),普通绘制更稳。
    $f.AllowTransparency = $true
    $f.Opacity = 0.94
    $f.BackColor = [System.Drawing.Color]::FromArgb(18, 20, 30)
    $f.add_Paint({
            param($sender, $e)
            $c = $script:PetOverlayContent
            if (-not $c) { return }
            $g = $e.Graphics
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
            if ($c.Kind -eq 'meme' -and $c.Meme) {
                $g.DrawImage($c.Meme, 0, 0, $sender.ClientSize.Width, $sender.ClientSize.Height)
                return
            }
            $fs = $(if ($script:PetOverlayFontSize) { $script:PetOverlayFontSize } else { 9.5 })
            $lh = $(if ($script:PetOverlayLineH) { $script:PetOverlayLineH } else { 22 })
            $pd = $(if ($script:PetOverlayPad) { $script:PetOverlayPad } else { 10 })
            # 换行/截断在 Set-PetOverlayContent 里已经算好,这里照着逐行画就行
            $font = New-Object System.Drawing.Font -ArgumentList 'Microsoft YaHei UI', $fs
            $y = $pd
            foreach ($ln in $c.Lines) {
                $g.DrawString([string]$ln, $font, [System.Drawing.Brushes]::White, $pd, $y)
                $y += $lh
            }
            foreach ($b in $c.Buttons) {
                $br = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(70, 90, 140, 230))
                $g.FillRectangle($br, $b.X, $b.Y, $b.W, $b.H)
                $br.Dispose()
                $cx = $b.X + $b.W / 2
                $cy = $b.Y + $b.H / 2
                $wb = [System.Drawing.Brushes]::White
                switch ($b.Glyph) {
                    'play' {
                        $pts = @(
                            (New-Object System.Drawing.PointF -ArgumentList ($cx - 5), ($cy - 7)),
                            (New-Object System.Drawing.PointF -ArgumentList ($cx + 7), $cy),
                            (New-Object System.Drawing.PointF -ArgumentList ($cx - 5), ($cy + 7))
                        )
                        $g.FillPolygon($wb, $pts)
                    }
                    'pause' {
                        $g.FillRectangle($wb, [int]($cx - 6), [int]($cy - 7), 4, 14)
                        $g.FillRectangle($wb, [int]($cx + 2), [int]($cy - 7), 4, 14)
                    }
                    'next' {
                        $g.FillRectangle($wb, [int]($cx + 5), [int]($cy - 7), 3, 14)
                        $pts = @(
                            (New-Object System.Drawing.PointF -ArgumentList ($cx - 6), ($cy - 7)),
                            (New-Object System.Drawing.PointF -ArgumentList ($cx - 6), ($cy + 7)),
                            (New-Object System.Drawing.PointF -ArgumentList ($cx + 4), $cy)
                        )
                        $g.FillPolygon($wb, $pts)
                    }
                    default {
                        $g.FillRectangle($wb, [int]($cx - 8), [int]($cy - 7), 3, 14)
                        $pts = @(
                            (New-Object System.Drawing.PointF -ArgumentList ($cx + 7), ($cy - 7)),
                            (New-Object System.Drawing.PointF -ArgumentList ($cx + 7), ($cy + 7)),
                            (New-Object System.Drawing.PointF -ArgumentList ($cx - 4), $cy)
                        )
                        $g.FillPolygon($wb, $pts)
                    }
                }
            }
            $font.Dispose()
        })
    $f.add_MouseEnter({ $script:PetOverlayHover = $true; $script:PetHoverOffAt = [datetime]::MinValue })
    $f.add_MouseLeave({ $script:PetOverlayHover = $false; $script:PetHoverOffAt = (Get-Date) })
    $f.add_MouseClick({
            param($s, $e)
            $c = $script:PetOverlayContent
            if (-not $c) { return }
            foreach ($b in $c.Buttons) {
                if ($e.X -ge $b.X -and $e.X -le ($b.X + $b.W) -and $e.Y -ge $b.Y -and $e.Y -le ($b.Y + $b.H)) {
                    if ($b.Action -eq 'meme') { Hide-PetOverlay }
                    else { [void](Send-PetMediaCommand -Action $b.Action) }
                    return
                }
            }
            if ($c.Kind -eq 'meme') { Hide-PetOverlay }
        })
    $script:PetOverlay = $f
    $f.Show()
    [CampusNetPet.Layered]::Enable($f.Handle)
    $f.Visible = $false
}

function Get-WrappedText {
    # 按像素宽度把一行拆成多行(中文逐字拆,够用且简单)
    param([string]$Text, $Font, [int]$MaxWidth, $Graphics)
    $out = @()
    if ([string]::IsNullOrEmpty($Text)) { return $out }
    if ($MaxWidth -le 20 -or -not $Graphics) { return @($Text) }
    $cur = ''
    foreach ($ch in $Text.ToCharArray()) {
        $try = $cur + $ch
        if ($cur -and ($Graphics.MeasureString($try, $Font)).Width -gt $MaxWidth) {
            # 别把英文/数字词从中间劈开:能退到最近的空格就退
            if ($ch -match '[0-9A-Za-z]') {
                $sp = $cur.LastIndexOf(' ')
                if ($sp -gt 0) {
                    $out += $cur.Substring(0, $sp)
                    $cur = $cur.Substring($sp + 1) + $ch
                    continue
                }
            }
            $out += $cur
            $cur = [string]$ch
        } else {
            $cur = $try
        }
    }
    if ($cur) { $out += $cur }
    return $out
}

function Set-PetOverlayContent {
    # Kind: info(网速+天气+音乐) / meme(表情包) / say(碎碎念气泡)
    # 尺寸:settings 里 petOverlayW / petOverlayH 给了就固定(文本自动换行 + 超长截断),给 0 就按内容自适应
    param([string]$Kind, $Meme = $null, [string]$Text = '')

    if (-not $script:PetOverlay) { return }
    if ($script:PetOverlayKind -eq 'meme' -and $Kind -eq 'meme') { return }

    $rawLines = New-Object System.Collections.ArrayList
    $buttons = @()
    $w = 250
    $h = 0
    $cfg = Get-PetCfg
    $fontSize = [double]$cfg.OverlayFont
    $pad = [int][Math]::Round($fontSize * 1.15)
    $lineH = [int][Math]::Ceiling($fontSize * 2.3)
    $fixedW = [int]$cfg.OverlayW
    $fixedH = [int]$cfg.OverlayH

    if ($Kind -eq 'meme' -and $Meme) {
        $w = [Math]::Min(300, $Meme.Width)
        $h = [int]([Math]::Round($Meme.Height * $w / $Meme.Width))
        $buttons = @(@{ X = 0; Y = 0; W = $w; H = $h; Action = 'meme'; Glyph = '' })
        $script:PetOverlay.Opacity = 1.0
    } else {
        $media = $null
        if ($Kind -eq 'say') {
            [void]$rawLines.Add($Text)
        } else {
            $sp = Get-NetSpeed
            if (-not $sp) { $sp = [pscustomobject]@{ DownText = '—'; UpText = '—' } }
            [void]$rawLines.Add('↓ ' + $sp.DownText + '   ↑ ' + $sp.UpText)
            $wx = Format-PetWeather
            if ($wx) { [void]$rawLines.Add('☁ ' + $wx) }
            $media = Get-PetMedia
            if ($media -and $media.Ok) { [void]$rawLines.Add('♪ ' + (Format-PetMediaLine)) }
        }

        # 量一遍:固定宽度就换行,固定高度就截断
        $measureFont = New-Object System.Drawing.Font -ArgumentList 'Microsoft YaHei UI', $fontSize
        $mb = New-Object System.Drawing.Bitmap -ArgumentList 1, 1
        $mg = [System.Drawing.Graphics]::FromImage($mb)

        $lines = New-Object System.Collections.ArrayList
        $textW = 0
        if ($fixedW -gt 0) {
            $maxText = $fixedW - $pad * 2 - 4
            foreach ($ln in $rawLines) {
                foreach ($seg in (Get-WrappedText -Text ([string]$ln) -Font $measureFont -MaxWidth $maxText -Graphics $mg)) {
                    [void]$lines.Add($seg)
                    $sw = ($mg.MeasureString($seg, $measureFont)).Width
                    if ($sw -gt $textW) { $textW = $sw }
                }
            }
        } else {
            foreach ($ln in $rawLines) {
                [void]$lines.Add([string]$ln)
                $sw = ($mg.MeasureString([string]$ln, $measureFont)).Width
                if ($sw -gt $textW) { $textW = $sw }
            }
        }
        $mg.Dispose(); $mb.Dispose(); $measureFont.Dispose()

        # 媒体按钮占的高度
        $btnH = 0
        if ($media -and $media.Ok -and $Kind -ne 'say') { $btnH = 36 }

        if ($fixedH -gt 0) {
            # 高度固定:能放几行放几行,放不下的截断并加省略号
            $avail = $fixedH - $pad * 2 - $btnH
            $maxLines = [Math]::Max(1, [int][Math]::Floor($avail / $lineH))
            if ($lines.Count -gt $maxLines) {
                $keep = @()
                for ($i = 0; $i -lt $maxLines; $i++) { $keep += [string]$lines[$i] }
                $last = $keep[$maxLines - 1]
                if ($last.Length -gt 1) { $keep[$maxLines - 1] = $last.Substring(0, $last.Length - 1) + '…' }
                $lines = New-Object System.Collections.ArrayList
                foreach ($k in $keep) { [void]$lines.Add($k) }
            }
        }

        if ($fixedW -gt 0) {
            $w = $fixedW
        } else {
            $w = [int][Math]::Ceiling($textW) + $pad * 2 + 6
            if ($media -and $media.Ok -and $Kind -ne 'say') {
                $need = $pad * 2 + 54 * 3 + 10 * 2
                if ($w -lt $need) { $w = $need }
            }
            $minW = $(if ($Kind -eq 'say') { 90 } else { 150 })
            if ($w -lt $minW) { $w = $minW }
        }

        if ($fixedH -gt 0) {
            $h = $fixedH
        } else {
            $h = $pad * 2 + $lines.Count * $lineH + $btnH
            if ($Kind -eq 'say') { $h = $pad * 2 + $lineH }
        }

        $script:PetOverlayFontSize = $fontSize
        $script:PetOverlayLineH = $lineH
        $script:PetOverlayPad = $pad
        $script:PetOverlayText = $Text

        if ($Kind -eq 'say') {
            $buttons = @(@{ X = 0; Y = 0; W = $w; H = $h; Action = 'meme'; Glyph = '' })
        } elseif ($media -and $media.Ok) {
            $bw = 54; $bh = 26; $gap = 10; $bx = [int]$pad
            $by = $h - $pad - $bh
            if ($by -lt 0) { $by = 0 }
            foreach ($it in @(
                    @{ Action = 'prev'; Glyph = 'prev' },
                    @{ Action = 'playpause'; Glyph = $(if ($media.Playing) { 'pause' } else { 'play' }) },
                    @{ Action = 'next'; Glyph = 'next' }
                )) {
                $buttons += @{ X = $bx; Y = $by; W = $bw; H = $bh; Action = $it.Action; Glyph = $it.Glyph }
                $bx += $bw + $gap
            }
        }
        $script:PetOverlay.Opacity = [double]$cfg.OverlayOpacity
        $rawLines = $lines
    }

    $script:PetOverlayContent = [pscustomobject]@{
        Kind    = $Kind
        Lines   = @($rawLines)
        Buttons = $buttons
        Meme    = $Meme
        Text    = $Text
    }
    $script:PetOverlayKind = $Kind
    $script:PetOverlay.ClientSize = New-Object System.Drawing.Size -ArgumentList $w, $h
    # 圆角:用 Region 裁窗口(直角会被切掉),信息框和碎碎念都适用
    if ($Kind -ne 'meme') {
        try {
            $r = 12
            $rp = New-Object System.Drawing.Drawing2D.GraphicsPath
            $rp.AddArc(0, 0, $r * 2, $r * 2, 180, 90)
            $rp.AddArc($w - $r * 2, 0, $r * 2, $r * 2, 270, 90)
            $rp.AddArc($w - $r * 2, $h - $r * 2, $r * 2, $r * 2, 0, 90)
            $rp.AddArc(0, $h - $r * 2, $r * 2, $r * 2, 90, 90)
            $rp.CloseFigure()
            $script:PetOverlay.Region = New-Object System.Drawing.Region -ArgumentList $rp
            $rp.Dispose()
        } catch { }
    } else {
        $script:PetOverlay.Region = $null
    }
    $script:PetOverlay.Visible = $true
    $script:PetOverlay.Invalidate()
    Update-PetOverlayPosition
}

function Update-PetOverlayPosition {
    if (-not $script:PetOverlay -or -not $script:PetOverlay.Visible) { return }
    $w = $script:PetOverlay.ClientSize.Width
    $h = $script:PetOverlay.ClientSize.Height
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $x = $script:PetX + 10
    $y = $script:PetY - $h - 6
    if ($y -lt $wa.Top) { $y = $script:PetY + 210 + 6 }
    if ($x + $w -gt $wa.Right) { $x = $wa.Right - $w - 4 }
    if ($x -lt $wa.Left) { $x = $wa.Left + 4 }
    $script:PetOverlay.Location = New-Object System.Drawing.Point -ArgumentList $x, $y
}

function Hide-PetOverlay {
    if ($script:PetOverlay -and -not $script:PetOverlay.IsDisposed) { $script:PetOverlay.Visible = $false }
    $script:PetOverlayKind = ''
    $script:PetOverlayContent = $null
}
# ---------------- 动作 / 走动 ----------------
function Get-PetCurrentAnim {
    if ($script:PetAction -and (Get-Date) -lt $script:PetActionEnd) { return $script:PetAction }
    if ($script:PetAction) { $script:PetAction = '' }
    if ($script:PetTarget) { return 'walk' }
    if (Test-PetShouldSleep) { return 'sleep' }
    return 'idle'
}

function Play-PetAction {
    param([string]$Id, [int]$HoldSec = 0)
    $dir = Join-Path $script:PetDir $Id
    if (-not (Test-Path $dir)) { return }
    $entry = Get-PetFrames -Anim $Id -Scale (Get-PetCfg).Scale
    if (-not $entry) { return }
    $fps = $entry.Fps
    $script:PetAction = $Id
    $dur = $(if ($HoldSec -gt 0) { $HoldSec } else { ($entry.Count / $fps) })
    $script:PetActionEnd = (Get-Date).AddSeconds($dur + 0.15)
    $script:PetAnim = $Id
    $script:PetLastIdx = -1
    if ($script:PetSw) { $script:PetSw.Restart() }
    $script:PetStillSince = (Get-Date).AddSeconds($dur)
    Write-Log ('宠物动作:' + (Get-PetLabel $Id))
}

function Start-PetWander {
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $tx = $script:PetXf + (Get-Random -Minimum -260 -Maximum 260)
    $ty = $script:PetYf + (Get-Random -Minimum -140 -Maximum 140)
    $tx = [Math]::Max($wa.Left + 10, [Math]::Min($wa.Right - 190, $tx))
    $ty = [Math]::Max($wa.Top + 10, [Math]::Min($wa.Bottom - 210, $ty))
    $script:PetTarget = @{ X = [double]$tx; Y = [double]$ty }
}

function Step-PetWander {
    if (-not $script:PetTarget) { return }
    $dx = $script:PetTarget.X - $script:PetXf
    $dy = $script:PetTarget.Y - $script:PetYf
    $d = [Math]::Sqrt($dx * $dx + $dy * $dy)
    if ($d -le 2) {
        $script:PetXf = $script:PetTarget.X
        $script:PetYf = $script:PetTarget.Y
        $script:PetTarget = $null
        $script:PetStillSince = Get-Date
        return
    }
    # 浮点累加,每秒约 96px;取整只在绘制时做,不会有整数截断造成的抖动
    $step = 1.6
    $script:PetXf = $script:PetXf + $dx / $d * $step
    $script:PetYf = $script:PetYf + $dy / $d * $step
}

# ---------------- 物理(拖拽抛掷) ----------------
function Start-PetPhysics {
    param([double]$Vx, [double]$Vy)
    $cfg = Get-PetCfg
    $script:PetVx = $Vx * $cfg.ThrowPower
    $script:PetVy = $Vy * $cfg.ThrowPower
    # 限制一下,免得一甩就飞出屏幕外找不回来
    if ([Math]::Abs($script:PetVx) -gt 4000) { $script:PetVx = 4000 * [Math]::Sign($script:PetVx) }
    if ([Math]::Abs($script:PetVy) -gt 4000) { $script:PetVy = 4000 * [Math]::Sign($script:PetVy) }
    $script:PetTarget = $null
    $script:PetPhys = ($script:PetVx -ne 0 -or $script:PetVy -ne 0)
    $script:PetLastPhysAt = Get-Date
    Write-Log ('宠物被甩出去: vx=' + [int]$script:PetVx + ' vy=' + [int]$script:PetVy)
}

function Step-PetPhysics {
    param([double]$Dt)
    if (-not $script:PetPhys) { return }
    if ($Dt -le 0) { $Dt = 0.016 }
    if ($Dt -gt 0.1) { $Dt = 0.1 }

    $cfg = Get-PetCfg
    $g = [double]$cfg.Gravity
    $e = [double]$cfg.Restitution
    $f = [double]$cfg.Friction

    $script:PetVy = $script:PetVy + $g * $Dt
    $script:PetXf = $script:PetXf + $script:PetVx * $Dt
    $script:PetYf = $script:PetYf + $script:PetVy * $Dt

    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $pw = $script:PetForm.ClientSize.Width
    $ph = $script:PetForm.ClientSize.Height
    $minX = $wa.Left
    $maxX = $wa.Right - $pw
    $minY = $wa.Top
    $maxY = $wa.Bottom - $ph

    $onGround = $false
    if ($script:PetYf -ge $maxY) {
        $script:PetYf = $maxY
        $onGround = $true
        if ([Math]::Abs($script:PetVy) -gt 60) { $script:PetVy = -$script:PetVy * $e }
        else { $script:PetVy = 0 }
        # 地面摩擦
        $script:PetVx = $script:PetVx * [Math]::Max(0, 1 - $f * $Dt)
        if ([Math]::Abs($script:PetVx) -lt 8) { $script:PetVx = 0 }
    }
    if ($cfg.CeilingBounce -and $script:PetYf -le $minY) {
        $script:PetYf = $minY
        if ([Math]::Abs($script:PetVy) -gt 60) { $script:PetVy = -$script:PetVy * $e } else { $script:PetVy = 0 }
    }
    if ($script:PetXf -le $minX) { $script:PetXf = $minX; $script:PetVx = -$script:PetVx * $e }
    if ($script:PetXf -ge $maxX) { $script:PetXf = $maxX; $script:PetVx = -$script:PetVx * $e }

    if ($onGround -and $script:PetVx -eq 0 -and $script:PetVy -eq 0) {
        $script:PetPhys = $false
        $script:PetStillSince = Get-Date
    }
}

# ---------------- 碎碎念 ----------------
function Get-PetSayings {
    if ($script:PetSayings) { return $script:PetSayings }
    $f = Join-Path $script:PetDir 'sayings.txt'
    $list = @()
    if (Test-Path $f) {
        try {
            $list = @(Get-Content $f -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        } catch { }
    }
    if ($list.Count -eq 0) { $list = @('……', '在的', '嗯?') }
    $script:PetSayings = $list
    return $list
}

function Show-PetSaying {
    param([string]$Text)
    $list = Get-PetSayings
    if (-not $Text) { $Text = [string]$list[(Get-Random -Maximum $list.Count)] }

    # 气泡时长:设了固定秒数就用它,设 0 就每次随机 2~6 秒
    $dur = 0.0
    try { $dur = [double](Get-PetCfg).SayDur } catch { }
    if ($dur -le 0) { $dur = [Math]::Round((Get-Random -Minimum 2.0 -Maximum 6.0), 1) }
    $script:PetSayUntil = (Get-Date).AddSeconds($dur)
    Set-PetOverlayContent -Kind 'say' -Meme $null -Text $Text
}

# ---------------- 主循环 ----------------
function Update-Pet {
    if (-not $script:PetForm -or $script:PetForm.IsDisposed) { return }
    if (-not $script:PetSw) { return }
    $script:PetTickAll++
    $now = Get-Date
    $cfg = Get-PetCfg

    # 缩放变了:强制重画一帧(缓存 key 变了,尺寸也会跟着变)
    if ([Math]::Abs($cfg.Scale - $script:PetLastScale) -gt 0.001) {
        $script:PetLastScale = $cfg.Scale
        $script:PetLastIdx = -1
        $script:PetLastDrawX = -99999
    }

    # 物理(拖拽甩出去之后的自由运动)
    if ($script:PetPhys) {
        $dtPhys = $(if ($script:PetLastPhysAt -ne [datetime]::MinValue) { ($now - $script:PetLastPhysAt).TotalSeconds } else { 0.016 })
        $script:PetLastPhysAt = $now
        Step-PetPhysics -Dt $dtPhys
    }

    # 碎碎念气泡到点收掉
    if ($script:PetSayUntil -ne [datetime]::MinValue -and $now -gt $script:PetSayUntil) {
        $script:PetSayUntil = [datetime]::MinValue
        if ($script:PetOverlayKind -eq 'say') { Hide-PetOverlay }
    }

    # 走动
    if ($script:PetTarget) { Step-PetWander }

    $anim = Get-PetCurrentAnim
    if ($anim -ne $script:PetAnim) {
        $script:PetAnim = $anim
        $script:PetLastIdx = -1
        $script:PetSw.Restart()
    }

    # 帧率变了(菜单切换 / 手改 settings.json)就重建帧缓存;不依赖菜单回调,自己会纠正
    if ([int]$cfg.Fps -ne $script:PetFpsApplied) {
        $script:PetFpsApplied = [int]$cfg.Fps
        foreach ($k in @($script:PetCache.Keys)) {
            foreach ($hb in $script:PetCache[$k].Hbs) { [CampusNetPet.Layered]::FreeHBitmap($hb) }
        }
        $script:PetCache = @{}
        $script:PetLastIdx = -1
        # 位图刚释放,催一下 GC,免得托管包装攒着(重建两次能涨几十 MB)
        [GC]::Collect()
        Write-Log ('宠物帧率:' + $script:PetFpsApplied + 'Hz(已重建帧缓存)')
    }

    $entry = Get-PetFrames -Anim $anim -Scale $cfg.Scale
    if (-not $entry) { return }
    $fps = $entry.Fps
    $idx = [int]([Math]::Floor($script:PetSw.Elapsed.TotalSeconds * $fps))
    if ($script:PetAction) {
        if ($idx -ge $entry.Count) { $idx = $entry.Count - 1 }
    } else {
        $idx = $idx % $entry.Count
    }

    $px = [int][Math]::Round($script:PetXf)
    $py = [int][Math]::Round($script:PetYf)
    $moved = ($px -ne $script:PetLastDrawX -or $py -ne $script:PetLastDrawY)

    # 位置变化也要重画:否则移动只能按动画帧率跳,看着就是"卡顿"
    if ($idx -ne $script:PetLastIdx -or $moved) {
        $script:PetLastIdx = $idx
        if ($script:PetForm.ClientSize.Width -ne $entry.W -or $script:PetForm.ClientSize.Height -ne $entry.H) {
            $script:PetForm.ClientSize = New-Object System.Drawing.Size -ArgumentList $entry.W, $entry.H
        }
        $null = [CampusNetPet.Layered]::Draw($script:PetForm.Handle, $entry.Hbs[$idx], $entry.W, $entry.H, $px, $py, [byte]255)
        $script:PetX = $px
        $script:PetY = $py
        $script:PetLastDrawX = $px
        $script:PetLastDrawY = $py
        $script:PetTicks++
    }

    # 面板跟着宠物走 / 内容刷新
    if ($script:PetOverlay -and $script:PetOverlay.Visible) { Update-PetOverlayPosition }

    if ($script:PetHover -and -not $script:PetAction) {
        if (($now - $script:PetOverlayRefreshAt).TotalMilliseconds -ge 500) {
            $script:PetOverlayRefreshAt = $now
            Set-PetOverlayContent -Kind 'info'
        }
    }

    # 悬停移出判定:宠物和面板都没在鼠标下才收
    if ($script:PetHoverOffAt -ne [datetime]::MinValue -and -not $script:PetHover -and -not $script:PetOverlayHover) {
        if (($now - $script:PetHoverOffAt).TotalMilliseconds -ge 350) {
            $script:PetHoverOffAt = [datetime]::MinValue
            Hide-PetOverlay
        }
    }

    # 表情包到点自动收
    if ($script:PetMemeUntil -ne [datetime]::MinValue -and $now -gt $script:PetMemeUntil) {
        $script:PetMemeUntil = [datetime]::MinValue
        Hide-PetOverlay
    }

    # 随机走动
    if (-not $script:PetTarget -and -not $script:PetAction) {
        $wanderOn = [bool]$cfg.Wander
        if ($wanderOn -and $script:PetStillSince -ne [datetime]::MinValue -and ($now - $script:PetStillSince).TotalSeconds -ge 45) {
            Start-PetWander
        }
    }

    # 随机自动碎碎念(和「随机动作间隔」分开:这个是随机时长随机触发)
    $sayEvery = [int]$cfg.SayEvery
    if ($cfg.Say -and $sayEvery -gt 0 -and -not $script:PetAction -and -not $script:PetTarget -and $now -gt $script:PetNextSay) {
        $lo = [Math]::Max(1, [int]($sayEvery * 0.5))
        $hi = [Math]::Max($lo + 1, [int]($sayEvery * 1.5))
        $script:PetNextSay = $now.AddMinutes((Get-Random -Minimum $lo -Maximum ($hi + 1)))
        if (-not ($script:PetOverlay -and $script:PetOverlay.Visible)) { Show-PetSaying }
    }

    # 随机动作 / 表情包
    $randMin = [int]$cfg.RandomMin
    if ($randMin -gt 0 -and $now -gt $script:PetNextRandom -and -not $script:PetAction -and -not $script:PetTarget) {
        $script:PetNextRandom = $now.AddMinutes($randMin)
        $pool = @(Get-PetAnimList | Where-Object { $_.id -ne 'idle' -and $_.id -ne 'sleep' -and $_.id -ne 'walk' })
        $memes = @(Get-PetMemeList)
        $roll = Get-Random -Minimum 0 -Maximum 100
        if ($memes.Count -gt 0 -and $roll -lt 30) {
            Show-PetMeme -Id ([string]$memes[(Get-Random -Maximum $memes.Count)].id)
        } elseif ([bool]$cfg.Say -and $roll -lt 55) {
            Show-PetSaying
        } elseif ($pool.Count -gt 0) {
            Play-PetAction -Id ([string]$pool[(Get-Random -Maximum $pool.Count)].id)
        }
    }

    # 帧率采样
    if ($script:PetTickStart -ne [datetime]::MinValue) {
        $el = ($now - $script:PetTickStart).TotalSeconds
        if (-not $script:PetFpsLog -and $el -ge 5) {
            $script:PetFpsLog = $true
            $script:PetTicksWarm = $script:PetTicks
            $script:PetWarmAt = $now
            Write-Log ('宠物渲染实测 ' + [Math]::Round($script:PetTicks / $el, 1) + ' fps / 计时器 ' + [Math]::Round($script:PetTickAll / $el, 1) + ' Hz(目标 ' + $fps + ')')
        } elseif ($script:PetFpsLog -and -not $script:PetFpsLog2 -and $script:PetWarmAt -ne [datetime]::MinValue -and ($now - $script:PetWarmAt).TotalSeconds -ge 10) {
            $script:PetFpsLog2 = $true
            $el2 = ($now - $script:PetWarmAt).TotalSeconds
            Write-Log ('宠物渲染稳态 ' + [Math]::Round(($script:PetTicks - $script:PetTicksWarm) / $el2, 1) + ' fps(目标 ' + $fps + ')')
        }
    }
}

function Save-PetPos {
    try {
        $s = Get-Settings
        $s.petX = [int]$script:PetX
        $s.petY = [int]$script:PetY
        Save-Settings -Settings $s
    } catch { }
}

# ---------- 宠物菜单的勾选:菜单只建一次,勾必须能跟着设置重算 ----------
# 以前是「建菜单时算一次」,所以把帧率改成 48Hz 再打开菜单,勾还停在 60Hz 上。
function Set-PetMenuCheck {
    # Name = 设置名,Value = 这一项的取值;petOverlayWH 是「宽,高」的组合项
    param($Item, [string]$Name, $Value, [double]$Tol = 0.01, $Snapshot)
    $Item | Add-Member -NotePropertyName PetCheckName -NotePropertyValue $Name -Force
    $Item | Add-Member -NotePropertyName PetCheckValue -NotePropertyValue $Value -Force
    $Item | Add-Member -NotePropertyName PetCheckTol -NotePropertyValue $Tol -Force
    Sync-PetMenuChecks -Items @($Item) -Snapshot $Snapshot
}

function Sync-PetMenuChecks {
    param($Items, $Snapshot)
    if ($null -eq $Snapshot) { try { $Snapshot = Get-Settings } catch { $Snapshot = $null } }
    foreach ($it in @($Items)) {
        $was = $false
        try { $was = [bool]$it.Checked } catch { }
        try {
            if ($Snapshot -and $it.PSObject.Properties['PetCheckName']) {
                $nm = [string]$it.PetCheckName
                if ($nm -eq 'petOverlayWH') {
                    $p2 = @(([string]$it.PetCheckValue) -split ',')
                    $it.Checked = ([int]$Snapshot.petOverlayW -eq [int]$p2[0] -and [int]$Snapshot.petOverlayH -eq [int]$p2[1])
                } else {
                    $cur = $it.PetCheckValue
                    if ($Snapshot.PSObject.Properties[$nm]) { $cur = $Snapshot.PSObject.Properties[$nm].Value }
                    # Tol < 0 表示开关类(直接按真假)
                    if ([double]$it.PetCheckTol -lt 0) { $it.Checked = [bool]$cur }
                    else { $it.Checked = ([Math]::Abs([double]$cur - [double]$it.PetCheckValue) -le [double]$it.PetCheckTol) }
                }
            }
        } catch { }
        try { if ([bool]$it.Checked -ne $was) { $script:PetCheckChanged = [int]$script:PetCheckChanged + 1 } } catch { }
        if ($it -is [System.Windows.Forms.ToolStripMenuItem] -and $it.DropDownItems.Count -gt 0) {
            Sync-PetMenuChecks -Items $it.DropDownItems -Snapshot $Snapshot
        }
    }
}

function Test-PetMenuPoint {
    # 鼠标点是不是落在菜单里(含所有还开着的二级 / 三级菜单)
    param($Menu, [System.Drawing.Point]$P)
    if (-not $Menu.Visible) { return $false }
    if ($Menu.Bounds.Contains($P)) { return $true }
    foreach ($it in @($Menu.Items)) {
        if ($it -is [System.Windows.Forms.ToolStripDropDownItem] -and $it.DropDown.Visible) {
            if (Test-PetMenuPoint -Menu $it.DropDown -P $P) { return $true }
        }
    }
    return $false
}

function New-PetMenu {
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    # 勾选状态统一用这一份快照,免得每个选项都去读一遍 settings.json
    $snap = $null
    try { $snap = Get-Settings } catch { }

    $miActions = New-Object System.Windows.Forms.ToolStripMenuItem
    $miActions.Text = '动作'
    foreach ($a in (Get-PetAnimList)) {
        if ($a.id -in @('idle', 'sleep', 'walk')) { continue }
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$a.label
        $mi.Tag = [string]$a.id
        $mi.add_Click({ param($sender, $e) Play-PetAction -Id ([string]$sender.Tag) })
        [void]$miActions.DropDownItems.Add($mi)
    }
    $miRandomNow = New-Object System.Windows.Forms.ToolStripMenuItem
    $miRandomNow.Text = '随机来一个'
    $miRandomNow.add_Click({
            $pool = @(Get-PetAnimList | Where-Object { $_.id -notin @('idle', 'sleep', 'walk') })
            if ($pool.Count -gt 0) { Play-PetAction -Id ([string]$pool[(Get-Random -Maximum $pool.Count)].id) }
        })
    $miMeme = New-Object System.Windows.Forms.ToolStripMenuItem
    $miMeme.Text = '随机表情包'
    $miMeme.add_Click({
            $memes = @(Get-PetMemeList)
            if ($memes.Count -gt 0) { Show-PetMeme -Id ([string]$memes[(Get-Random -Maximum $memes.Count)].id) }
        })
    [void]$miActions.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$miActions.DropDownItems.Add($miRandomNow)
    [void]$miActions.DropDownItems.Add($miMeme)
    [void]$miActions.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miInterval = New-Object System.Windows.Forms.ToolStripMenuItem
    $miInterval.Text = '随机间隔'
    $cur = [int](Get-PetSetting -Name 'petRandomMin' -Default 5)
    foreach ($opt in @(@{ m = 0; t = '关' }, @{ m = 1; t = '每 1 分钟' }, @{ m = 5; t = '每 5 分钟' }, @{ m = 15; t = '每 15 分钟' }, @{ m = 30; t = '每 30 分钟' })) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petRandomMin' -Value ([int]$opt.m) -Tol 0 -Snapshot $snap
        $mi.Tag = [int]$opt.m
        $mi.add_Click({
                param($sender, $e)
                Set-PetSettingValue -Name 'petRandomMin' -Value ([int]$sender.Tag)
                $script:PetCfg = $null; $script:PetCfgAt = [datetime]::MinValue
                $script:PetNextRandom = (Get-Date).AddMinutes([int]$sender.Tag)
                Write-Log ('宠物随机动作间隔改为: ' + $sender.Text)
            })
        [void]$miInterval.DropDownItems.Add($mi)
    }
    [void]$miActions.DropDownItems.Add($miInterval)
    [void]$menu.Items.Add($miActions)


    $miWander = New-Object System.Windows.Forms.ToolStripMenuItem
    $miWander.Text = '随机走动'
    $miWander.CheckOnClick = $true
    Set-PetMenuCheck -Item $miWander -Name 'petWander' -Value $true -Tol -1 -Snapshot $snap
    $miWander.add_Click({
            param($sender, $e)
            Set-PetSettingValue -Name 'petWander' -Value ([bool]$sender.Checked)
            if (-not $sender.Checked) { $script:PetTarget = $null }
        })
    [void]$menu.Items.Add($miWander)

    $miWeather = New-Object System.Windows.Forms.ToolStripMenuItem
    $miWeather.Text = '悬停显示天气(联网定位)'
    $miWeather.CheckOnClick = $true
    Set-PetMenuCheck -Item $miWeather -Name 'petWeather' -Value $true -Tol -1 -Snapshot $snap
    $miWeather.add_Click({
            param($sender, $e)
            Set-PetSettingValue -Name 'petWeather' -Value ([bool]$sender.Checked)
            if ($sender.Checked) { Start-PetWeatherRefresh -Force }
        })
    [void]$menu.Items.Add($miWeather)

    $miSize = New-Object System.Windows.Forms.ToolStripMenuItem
    $miSize.Text = '大小'
    $curSize = [double](Get-PetSetting -Name 'petScale' -Default 1.0)
    foreach ($opt in @(0.6, 0.8, 1.0, 1.25, 1.5, 2.0)) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = ('{0}%' -f [int][Math]::Round($opt * 100))
        Set-PetMenuCheck -Item $mi -Name 'petScale' -Value ([double]$opt) -Tol 0.01 -Snapshot $snap
        $mi.Tag = [double]$opt
        $mi.add_Click({
                param($sender, $e)
                Set-PetSize -Scale ([double]$sender.Tag)
            })
        [void]$miSize.DropDownItems.Add($mi)
    }
    [void]$menu.Items.Add($miSize)

    $miBox = New-Object System.Windows.Forms.ToolStripMenuItem
    $miBox.Text = '信息框'
    $curFont = [double](Get-PetSetting -Name 'petOverlayFont' -Default 9.5)
    $miFont = New-Object System.Windows.Forms.ToolStripMenuItem
    $miFont.Text = '字号'
    foreach ($opt in @(@{ v = 8; t = '小' }, @{ v = 9.5; t = '中' }, @{ v = 11; t = '大' }, @{ v = 13; t = '特大' })) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petOverlayFont' -Value ([double]$opt.v) -Tol 0.01 -Snapshot $snap
        $mi.Tag = [double]$opt.v
        $mi.add_Click({
                param($sender, $e)
                Set-PetSettingValue -Name 'petOverlayFont' -Value ([double]$sender.Tag)
                $script:PetCfg = $null
                $script:PetCfgAt = [datetime]::MinValue
                if ($script:PetHover) { Set-PetOverlayContent -Kind 'info' }
            })
        [void]$miFont.DropDownItems.Add($mi)
    }
    [void]$miBox.DropDownItems.Add($miFont)

    $miAlpha = New-Object System.Windows.Forms.ToolStripMenuItem
    $miAlpha.Text = '透明度'
    $curAlpha = [double](Get-PetSetting -Name 'petOverlayOpacity' -Default 0.94)
    foreach ($opt in @(@{ v = 0.7; t = '70%' }, @{ v = 0.85; t = '85%' }, @{ v = 1.0; t = '不透明' })) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petOverlayOpacity' -Value ([double]$opt.v) -Tol 0.02 -Snapshot $snap
        $mi.Tag = [double]$opt.v
        $mi.add_Click({
                param($sender, $e)
                Set-PetSettingValue -Name 'petOverlayOpacity' -Value ([double]$sender.Tag)
                $script:PetCfg = $null
                $script:PetCfgAt = [datetime]::MinValue
                if ($script:PetHover) { Set-PetOverlayContent -Kind 'info' }
            })
        [void]$miAlpha.DropDownItems.Add($mi)
    }
    [void]$miBox.DropDownItems.Add($miAlpha)
    # 尺寸档位(0 = 自动)
    $curOW = [int](Get-PetSetting -Name 'petOverlayW' -Default 0)
    $curOH = [int](Get-PetSetting -Name 'petOverlayH' -Default 0)
    $miBoxSize = New-Object System.Windows.Forms.ToolStripMenuItem
    $miBoxSize.Text = '尺寸'
    foreach ($opt in @(
            @{ t = '自动(按内容)'; w = 0; h = 0 },
            @{ t = '紧凑 240 × 自动'; w = 240; h = 0 },
            @{ t = '标准 320 × 自动'; w = 320; h = 0 },
            @{ t = '宽版 420 × 自动'; w = 420; h = 0 },
            @{ t = '方框 300 × 120'; w = 300; h = 120 },
            @{ t = '大框 380 × 160'; w = 380; h = 160 }
        )) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petOverlayWH' -Value ([string]$opt.w + ',' + [string]$opt.h) -Tol 0 -Snapshot $snap
        $mi.Tag = ([string]$opt.w + ',' + [string]$opt.h)
        $mi.add_Click({
                param($sender, $e)
                $parts = @(([string]$sender.Tag) -split ',')
                Set-PetOverlaySize -W ([int]$parts[0]) -H ([int]$parts[1])
            })
        [void]$miBoxSize.DropDownItems.Add($mi)
    }
    [void]$miBoxSize.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miBoxCustom = New-Object System.Windows.Forms.ToolStripMenuItem
    $miBoxCustom.Text = '自定义尺寸…'
    $miBoxCustom.add_Click({
            try {
                Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
                $cw = [int](Get-PetSetting -Name 'petOverlayW' -Default 0)
                $ch = [int](Get-PetSetting -Name 'petOverlayH' -Default 0)
                $ans = [Microsoft.VisualBasic.Interaction]::InputBox('宽,高(像素;0 = 自动)。例:320,0 或 320,120', '信息框尺寸', ($cw.ToString() + ',' + $ch.ToString()))
                if ($null -ne $ans -and $ans.Trim()) {
                    $parts = @($ans -split '[,，xX* ]+' | Where-Object { $_ -ne '' })
                    $nw = 0; $nh = 0
                    if ($parts.Count -ge 1) { [void][int]::TryParse([string]$parts[0], [ref]$nw) }
                    if ($parts.Count -ge 2) { [void][int]::TryParse([string]$parts[1], [ref]$nh) }
                    Set-PetOverlaySize -W $nw -H $nh
                }
            } catch { }
        })
    [void]$miBoxSize.DropDownItems.Add($miBoxCustom)
    [void]$miBox.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$miBox.DropDownItems.Add($miBoxSize)
    [void]$menu.Items.Add($miBox)

    # 帧率(24 = 素材原生;48/60 = 插中间帧,动作时长不变)
    $miFps = New-Object System.Windows.Forms.ToolStripMenuItem
    $miFps.Text = '帧率'
    $curFps = [int](Get-PetSetting -Name 'petFps' -Default 60)
    foreach ($opt in @(
            @{ v = 24; t = '24Hz(原生,最省内存)' },
            @{ v = 48; t = '48Hz(插一帧)' },
            @{ v = 60; t = '60Hz(插一帧 + 60Hz 计时器)' }
        )) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petFps' -Value ([int]$opt.v) -Tol 0 -Snapshot $snap
        $mi.Tag = [int]$opt.v
        $mi.add_Click({
                param($sender, $e)
                $v = [int]$sender.Tag
                Set-PetSettingValue -Name 'petFps' -Value $v
                # 帧缓存是按旧帧率插值的,得清掉重来
                foreach ($k in @($script:PetCache.Keys)) {
                    foreach ($hb in $script:PetCache[$k].Hbs) { [CampusNetPet.Layered]::FreeHBitmap($hb) }
                }
                $script:PetCache = @{}
                $script:PetLastIdx = -1
                if ($script:PetTimer) { $script:PetTimer.Interval = 4 }
                Write-Log ('宠物帧率: ' + $v + 'Hz')
            })
        [void]$miFps.DropDownItems.Add($mi)
    }
    [void]$menu.Items.Add($miFps)
    # 物理(拖拽抛掷手感)
    $miPhys = New-Object System.Windows.Forms.ToolStripMenuItem
    $miPhys.Text = '物理'
    $cfgNow = Get-PetCfg
    $miFeel = New-Object System.Windows.Forms.ToolStripMenuItem
    $miFeel.Text = '手感'
    foreach ($opt in @(@{ g = 900; t = '轻盈' }, @{ g = 1400; t = '标准' }, @{ g = 2200; t = '厚重' })) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petGravity' -Value ([int]$opt.g) -Tol 1 -Snapshot $snap
        $mi.Tag = [int]$opt.g
        $mi.add_Click({
                param($sender, $e)
                Set-PetSettingValue -Name 'petGravity' -Value ([int]$sender.Tag)
                $script:PetCfg = $null; $script:PetCfgAt = [datetime]::MinValue
                Write-Log ('宠物物理手感: ' + $sender.Text)
            })
        [void]$miFeel.DropDownItems.Add($mi)
    }
    [void]$miPhys.DropDownItems.Add($miFeel)

    $miRest = New-Object System.Windows.Forms.ToolStripMenuItem
    $miRest.Text = '弹性'
    foreach ($opt in @(@{ v = 0.5; t = '低' }, @{ v = 0.78; t = '标准' }, @{ v = 0.95; t = '高' })) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petRestitution' -Value ([double]$opt.v) -Tol 0.02 -Snapshot $snap
        $mi.Tag = [double]$opt.v
        $mi.add_Click({
                param($sender, $e)
                Set-PetSettingValue -Name 'petRestitution' -Value ([double]$sender.Tag)
                $script:PetCfg = $null; $script:PetCfgAt = [datetime]::MinValue
            })
        [void]$miRest.DropDownItems.Add($mi)
    }
    [void]$miPhys.DropDownItems.Add($miRest)

    $miFric = New-Object System.Windows.Forms.ToolStripMenuItem
    $miFric.Text = '地面摩擦'
    foreach ($opt in @(@{ v = 1.0; t = '小' }, @{ v = 2.5; t = '标准' }, @{ v = 5.0; t = '大' })) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petGroundFriction' -Value ([double]$opt.v) -Tol 0.05 -Snapshot $snap
        $mi.Tag = [double]$opt.v
        $mi.add_Click({
                param($sender, $e)
                Set-PetSettingValue -Name 'petGroundFriction' -Value ([double]$sender.Tag)
                $script:PetCfg = $null; $script:PetCfgAt = [datetime]::MinValue
            })
        [void]$miFric.DropDownItems.Add($mi)
    }
    [void]$miPhys.DropDownItems.Add($miFric)

    $miCeil = New-Object System.Windows.Forms.ToolStripMenuItem
    $miCeil.Text = '顶部反弹'
    $miCeil.CheckOnClick = $true
    Set-PetMenuCheck -Item $miCeil -Name 'petCeilingBounce' -Value $true -Tol -1 -Snapshot $snap
    $miCeil.add_Click({
            param($sender, $e)
            Set-PetSettingValue -Name 'petCeilingBounce' -Value ([bool]$sender.Checked)
            $script:PetCfg = $null; $script:PetCfgAt = [datetime]::MinValue
        })
    [void]$miPhys.DropDownItems.Add($miCeil)
    [void]$menu.Items.Add($miPhys)

    # 碎碎念(开关 + 随机触发间隔 + 气泡时长)
    $miSay = New-Object System.Windows.Forms.ToolStripMenuItem
    $miSay.Text = '碎碎念'

    $miSayOn = New-Object System.Windows.Forms.ToolStripMenuItem
    $miSayOn.Text = '自动碎碎念'
    $miSayOn.CheckOnClick = $true
    Set-PetMenuCheck -Item $miSayOn -Name 'petSay' -Value $true -Tol -1 -Snapshot $snap
    $miSayOn.add_Click({
            param($sender, $e)
            Set-PetSettingValue -Name 'petSay' -Value ([bool]$sender.Checked)
            $script:PetCfg = $null; $script:PetCfgAt = [datetime]::MinValue
            $script:PetNextSay = (Get-Date)
        })
    [void]$miSay.DropDownItems.Add($miSayOn)

    $miSayEvery = New-Object System.Windows.Forms.ToolStripMenuItem
    $miSayEvery.Text = '随机触发'
    $curEvery = [int](Get-PetSetting -Name 'petSayEvery' -Default 6)
    foreach ($opt in @(
            @{ m = 0; t = '关(只跟随机动作一起)' },
            @{ m = 6; t = '频繁:每 3~10 分钟' },
            @{ m = 20; t = '一般:每 10~30 分钟' },
            @{ m = 60; t = '偶尔:每 30~90 分钟' }
        )) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petSayEvery' -Value ([int]$opt.m) -Tol 0 -Snapshot $snap
        $mi.Tag = [int]$opt.m
        $mi.add_Click({
                param($sender, $e)
                $v = [int]$sender.Tag
                Set-PetSettingValue -Name 'petSayEvery' -Value $v
                $script:PetCfg = $null; $script:PetCfgAt = [datetime]::MinValue
                $script:PetNextSay = (Get-Date)
                Write-Log ('碎碎念随机触发: ' + $sender.Text)
            })
        [void]$miSayEvery.DropDownItems.Add($mi)
    }
    [void]$miSay.DropDownItems.Add($miSayEvery)

    $miSayDur = New-Object System.Windows.Forms.ToolStripMenuItem
    $miSayDur.Text = '气泡时长'
    $curDur = [double](Get-PetSetting -Name 'petSayDur' -Default 0)
    foreach ($opt in @(
            @{ s = 0; t = '随机 2~6 秒' },
            @{ s = 3; t = '3 秒' },
            @{ s = 5; t = '5 秒' },
            @{ s = 8; t = '8 秒' }
        )) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem
        $mi.Text = [string]$opt.t
        Set-PetMenuCheck -Item $mi -Name 'petSayDur' -Value ([double]$opt.s) -Tol 0.01 -Snapshot $snap
        $mi.Tag = [double]$opt.s
        $mi.add_Click({
                param($sender, $e)
                Set-PetSettingValue -Name 'petSayDur' -Value ([double]$sender.Tag)
                $script:PetCfg = $null; $script:PetCfgAt = [datetime]::MinValue
                Write-Log ('碎碎念气泡时长: ' + $sender.Text)
            })
        [void]$miSayDur.DropDownItems.Add($mi)
    }
    [void]$miSay.DropDownItems.Add($miSayDur)

    [void]$miSay.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miSayNow = New-Object System.Windows.Forms.ToolStripMenuItem
    $miSayNow.Text = '说一句'
    $miSayNow.add_Click({ Show-PetSaying })
    [void]$miSay.DropDownItems.Add($miSayNow)
    [void]$menu.Items.Add($miSay)

    # 天气城市
    $miCity = New-Object System.Windows.Forms.ToolStripMenuItem
    $miCity.Text = '天气城市…'
    $miCity.add_Click({
            try {
                Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
                $cur = [string](Get-PetSetting -Name 'petWeatherCity' -Default '')
                $ans = [Microsoft.VisualBasic.Interaction]::InputBox('输入城市名(留空=按 IP 自动定位)', '天气城市', $cur)
                if ($null -ne $ans) {
                    Set-PetSettingValue -Name 'petWeatherCity' -Value ([string]$ans).Trim()
                    $script:PetWeather = $null
                    $script:PetWeatherAt = [datetime]::MinValue
                    Start-PetWeatherRefresh -Force
                    Write-Log ('天气城市: ' + $(if ($ans.Trim()) { $ans.Trim() } else { '(按 IP 定位)' }))
                }
            } catch { }
        })
    [void]$menu.Items.Add($miCity)


    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $miMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $miMenu.Text = '打开设置菜单'
    $miMenu.add_Click({ Start-Process -FilePath 'powershell.exe' -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $SelfScript) })
    [void]$menu.Items.Add($miMenu)

    $miAbout = New-Object System.Windows.Forms.ToolStripMenuItem
    $miAbout.Text = '关于 / 素材来源(dsh-pet)'
    $miAbout.add_Click({ Start-Process $script:PetUrl })
    [void]$menu.Items.Add($miAbout)

    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miHide = New-Object System.Windows.Forms.ToolStripMenuItem
    $miHide.Text = '隐藏宠物(托盘里可再打开)'
    $miHide.add_Click({ Stop-Pet -Remember })
    [void]$menu.Items.Add($miHide)

    return $menu
}

function Start-PetWeatherRefresh {
    param([switch]$Force)
    if (-not [bool](Get-PetSetting -Name 'petWeather' -Default $true)) { return }
    $fresh = $false
    if (Test-Path $script:PetWeatherFile) {
        try {
            $age = ((Get-Date) - (Get-Item $script:PetWeatherFile).LastWriteTime).TotalMinutes
            $fresh = ($age -lt 30)
        } catch { }
    }
    if ($fresh -and -not $Force) { return }
    try {
        Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList (
            '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Mode weather -Background' -f $SelfScript)
    } catch { }
}

function Start-Pet {
    param([switch]$Quiet)

    if ($script:PetForm -and -not $script:PetForm.IsDisposed) { return $true }
    if (-not (Test-PetAssets)) {
        Write-Log ('没有宠物素材(缺 ' + (Join-Path $script:PetDir 'idle') + '),跳过') 'WARN'
        return $false
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [void](Set-PetDpiAware)
    if (-not ('CampusNetPet.Layered' -as [type])) {
        Add-Type -TypeDefinition $script:LayeredCs -Language CSharp -ReferencedAssemblies System.Drawing
    }

    $frames = Get-PetFrames -Anim 'idle' -Scale ([double](Get-PetSetting -Name 'petScale' -Default 1.0))
    if (-not $frames) { Write-Log '宠物帧读不出来(pet/idle/*.png)', 'WARN'; return $false }
    $w = $frames.W
    $h = $frames.H

    $set = Get-Settings
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    if ([int]$set.petX -gt 0 -and [int]$set.petY -gt 0) {
        $script:PetX = [int]$set.petX
        $script:PetY = [int]$set.petY
    } else {
        $script:PetX = $wa.Left + 40
        $script:PetY = $wa.Bottom - $h - 60
    }
    if ($script:PetX -gt ($wa.Right - 40)) { $script:PetX = $wa.Right - $w - 30 }
    if ($script:PetX -lt $wa.Left) { $script:PetX = $wa.Left + 10 }
    if ($script:PetY -gt ($wa.Bottom - 40)) { $script:PetY = $wa.Bottom - $h - 40 }
    if ($script:PetY -lt $wa.Top) { $script:PetY = $wa.Top + 10 }

    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.ShowInTaskbar = $false
    $form.TopMost = $true
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $form.BackColor = [System.Drawing.Color]::FromArgb(1, 1, 1)
    $form.Text = '校园网宠物'
    $form.ClientSize = New-Object System.Drawing.Size -ArgumentList $w, $h
    $form.Location = New-Object System.Drawing.Point -ArgumentList $script:PetX, $script:PetY
    $script:PetMenu = New-PetMenu
    $form.ContextMenuStrip = $script:PetMenu
    # 菜单只在宠物启动时建一次 —— 每次弹出前按最新设置重算勾选,
    # 否则改成 48Hz 再打开菜单,勾还留在 60Hz 上。
    $script:PetMenu.add_Opening({
            try {
                $script:PetCheckChanged = 0
                Sync-PetMenuChecks -Items $script:PetMenu.Items
                if ([int]$script:PetCheckChanged -gt 0) {
                    Write-Log ('宠物菜单勾选已按最新设置重算(' + [int]$script:PetCheckChanged + ' 项变化)')
                }
            } catch { }
        })
    # 点菜单外面自动关闭(窗口 TopMost 且没有焦点时,默认行为常常不收起)
    # 关键:必须等「松手」再判断。按下就收会把二级菜单里的点击吃掉 ——
    # 帧率 / 大小 / 信息框 / 动作 / 随机间隔 / 碎碎念 全在二级菜单里,
    # 以前这些项点了完全没反应(日志里连一行「宠物帧率:48Hz」都没有)就是这个原因。
    $script:PetMenuGuard = New-Object System.Windows.Forms.Timer
    $script:PetMenuGuard.Interval = 40
    $script:PetMenuGuard.add_Tick({
            if (-not $script:PetMenu -or -not $script:PetMenu.Visible) { $script:PetMenuGuard.Stop(); $script:PetMenuDown = $false; return }
            $btn = [System.Windows.Forms.Control]::MouseButtons
            if ($btn -ne [System.Windows.Forms.MouseButtons]::None) { $script:PetMenuDown = $true; return }
            if (-not $script:PetMenuDown) { return }
            $script:PetMenuDown = $false
            $cp = [System.Windows.Forms.Cursor]::Position
            if (-not (Test-PetMenuPoint -Menu $script:PetMenu -P $cp)) {
                $script:PetMenu.Close()
                $script:PetMenuGuard.Stop()
            }
        })
    $script:PetMenu.add_Opened({ $script:PetMenuDown = $false; $script:PetMenuGuard.Start() })
    $script:PetMenu.add_Closed({ $script:PetMenuGuard.Stop() })
    $form.add_MouseEnter({
            $script:PetHover = $true
            $script:PetHoverOffAt = [datetime]::MinValue
            Set-PetOverlayContent -Kind 'info'
        })
    $form.add_MouseLeave({ $script:PetHover = $false; $script:PetHoverOffAt = (Get-Date) })
    $form.add_MouseDown({
            param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $script:PetDrag = @{ X = $e.X; Y = $e.Y; Moved = $false; VX = 0.0; VY = 0.0; At = (Get-Date); LastX = $script:PetXf; LastY = $script:PetYf }
                $script:PetTarget = $null
                $script:PetPhys = $false
                $script:PetVx = 0; $script:PetVy = 0
            }
        })
    $form.add_MouseMove({
            param($s, $e)
            if ($script:PetDrag) {
                $p = [System.Windows.Forms.Cursor]::Position
                $nx = [double]($p.X - $script:PetDrag.X)
                $ny = [double]($p.Y - $script:PetDrag.Y)
                $now = Get-Date
                $dt = ($now - $script:PetDrag.At).TotalSeconds
                if ($dt -gt 0.008) {
                    # 记录拖动速度(px/s),松手时当抛掷初速
                    $script:PetDrag.VX = ($nx - $script:PetDrag.LastX) / $dt
                    $script:PetDrag.VY = ($ny - $script:PetDrag.LastY) / $dt
                    $script:PetDrag.LastX = $nx
                    $script:PetDrag.LastY = $ny
                    $script:PetDrag.At = $now
                }
                $script:PetXf = $nx
                $script:PetYf = $ny
                $script:PetX = [int][Math]::Round($script:PetXf)
                $script:PetY = [int][Math]::Round($script:PetYf)
                $script:PetDrag.Moved = $true
                Update-Pet
            }
        })
    $form.add_MouseUp({
            param($s, $e)
            if ($script:PetDrag) {
                if ($script:PetDrag.Moved) {
                    Save-PetPos
                    # 甩出去:按松手瞬间的速度给一个初速,之后交给重力
                    $speed = [Math]::Sqrt($script:PetDrag.VX * $script:PetDrag.VX + $script:PetDrag.VY * $script:PetDrag.VY)
                    if ($speed -gt 120) { Start-PetPhysics -Vx $script:PetDrag.VX -Vy $script:PetDrag.VY }
                }
                else {
                    # 点一下 / 摸头:随机挑一个动作(上半部分算摸头)
                    $head = ($e.Y -lt ($s.ClientSize.Height * 0.4))
                    $pool = @(Get-PetAnimList | Where-Object { $_.id -notin @('idle', 'sleep', 'walk') })
                    $pick = 'wave'
                    if ($pool.Count -gt 0) { $pick = [string]$pool[(Get-Random -Maximum $pool.Count)].id }
                    Play-PetAction -Id $pick
                    Write-Log ('宠物被' + $(if ($head) { '摸头' } else { '点了一下' }) + ' → ' + (Get-PetLabel $pick))
                    if ([bool](Get-PetCfg).Say -and (Get-Random -Maximum 100) -lt 30) { Show-PetSaying }
                }
                $script:PetDrag = $null
            }
        })

    $script:PetForm = $form
    $script:PetAnim = ''
    $script:PetLastIdx = -1
    $script:PetAction = ''
    $script:PetTarget = $null
    $script:PetStillSince = Get-Date
    $script:PetXf = [double]$script:PetX
    $script:PetYf = [double]$script:PetY
    $script:PetLastDrawX = -99999
    $script:PetLastDrawY = -99999
    $script:PetCfg = $null
    $script:PetCfgAt = [datetime]::MinValue
    $script:PetNextRandom = (Get-Date).AddMinutes([int](Get-PetSetting -Name 'petRandomMin' -Default 5))
    $script:PetNextSay = (Get-Date)
    $script:PetSw = New-Object System.Diagnostics.Stopwatch
    $script:PetSw.Start()
    $script:PetTickStart = Get-Date
    $script:PetTicks = 0
    $script:PetTickAll = 0
    $script:PetFpsLog = $false
    $script:PetFpsLog2 = $false
    $script:PetTicksWarm = 0
    $script:PetWarmAt = [datetime]::MinValue

    $form.Show()
    [CampusNetPet.Layered]::Enable($form.Handle)
    [CampusNetPet.Layered]::BeginHighRes()
    $noThrottle = [CampusNetPet.Layered]::DisablePowerThrottling()
    Write-Log ('宠物计时器: 1ms 精度, 已' + $(if ($noThrottle) { '退出' } else { '未退出(可能被系统强制)' }) + '电源节流')

    New-PetOverlayWindow

    $timer = New-Object System.Windows.Forms.Timer
    $fpsWant = [int](Get-PetSetting -Name 'petFps' -Default 60)
    if ($fpsWant -lt 12) { $fpsWant = 24 }
    # WinForms 的 Timer 分辨率约 15.6ms,用 4ms 间隔让它跑满;实际节拍由每帧工作量决定
    $timer.Interval = 4
    $timer.add_Tick({ try { Update-Pet } catch { } })
    $timer.Start()
    $script:PetTimer = $timer

    Update-Pet
    Start-PetWeatherRefresh
    if (-not $Quiet) {
        $anims = @(Get-PetAnimList | Where-Object { $_.id -notin @('idle', 'sleep', 'walk') }).Count
        $memes = @(Get-PetMemeList).Count
        Write-Log ('桌面宠物已启动(' + $w + 'x' + $h + ',' + $fpsWant + 'Hz(素材 ' + (Get-PetFps) + 'fps),' + $anims + ' 个动作,' + $memes + ' 张表情包,素材:' + $script:PetUrl + ')') 'OK'
    }
    return $true
}

function Stop-Pet {
    param([switch]$Remember)

    $had = [bool]($script:PetForm -or $script:PetTimer)
    if ($script:PetTimer) { try { $script:PetTimer.Stop(); $script:PetTimer.Dispose() } catch { }; $script:PetTimer = $null }
    if ($script:PetOverlay -and -not $script:PetOverlay.IsDisposed) {
        try { $script:PetOverlay.Hide(); $script:PetOverlay.Dispose() } catch { }
    }
    $script:PetOverlay = $null
    if ($script:PetForm -and -not $script:PetForm.IsDisposed) {
        try { $script:PetForm.Hide(); $script:PetForm.Dispose() } catch { }
    }
    if ('CampusNetPet.Layered' -as [type]) { try { [CampusNetPet.Layered]::EndHighRes() } catch { } }
    $script:PetForm = $null
    if ($script:PetSw) { $script:PetSw.Stop(); $script:PetSw = $null }
    foreach ($k in @($script:PetCache.Keys)) {
        foreach ($hb in $script:PetCache[$k].Hbs) { [CampusNetPet.Layered]::FreeHBitmap($hb) }
    }
    $script:PetCache = @{}
    foreach ($k in @($script:PetMemeCache.Keys)) { try { $script:PetMemeCache[$k].Dispose() } catch { } }
    $script:PetMemeCache = @{}
    $script:PetAnim = ''
    $script:PetAction = ''
    $script:PetTarget = $null
    $script:PetLastIdx = -1

    if ($Remember) {
        try {
            $s = Get-Settings
            $s.pet = $false
            Save-Settings -Settings $s
        } catch { }
    }
    if ($had) { Write-Log '桌面宠物已关闭' }
}

function Set-PetEnabled {
    param([bool]$On)
    try {
        $s = Get-Settings
        $s.pet = $On
        Save-Settings -Settings $s
    } catch { }
    if ($On) { [void](Start-Pet) } else { Stop-Pet }
}

function Sync-PetWithSettings {
    $stamp = 0
    try { if (Test-Path $SettingsFn) { $stamp = (Get-Item $SettingsFn).LastWriteTimeUtc.Ticks } } catch { }
    if ($stamp -eq $script:PetSettingsStamp) { return }
    $script:PetSettingsStamp = $stamp

    $want = $false
    try { $want = [bool](Get-Settings).pet } catch { }
    $on = [bool]($script:PetForm -and -not $script:PetForm.IsDisposed)
    if ($want -and -not $on) { [void](Start-Pet) }
    elseif (-not $want -and $on) { Stop-Pet }
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
    # Get-CimInstance 查守护进程要几十毫秒,每 2 秒查一次会把宠物顶掉一两帧 —— 缓存 10 秒
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

    # 先声明 DPI 感知(必须在建窗口之前),否则宠物窗口在高缩放下会被放大并偏移
    [void](Set-PetDpiAware)

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

    $itemPet = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemPet.Text = '桌面宠物: ' + $(if ((Get-Settings).pet) { '开' } else { '关' })
    $itemPet.add_Click({
            $on = -not [bool](Get-Settings).pet
            Set-PetEnabled -On $on
            $itemPet.Text = '桌面宠物: ' + $(if ($on) { '开' } else { '关' })
        })
    [void]$menu.Items.Add($itemPet)

    $itemMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemMenu.Text = '打开设置菜单…'
    $itemMenu.add_Click({
            Start-Process -FilePath 'powershell.exe' `
                -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $SelfScript)
        })
    [void]$menu.Items.Add($itemMenu)

    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $itemExit = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemExit.Text = '退出托盘(同时停止守护和桌宠)'
    $itemExit.add_Click({
            Stop-Watcher
            Stop-PetProcess
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

    # 宠物跟着托盘一起起来(开关在托盘右键 / 菜单 P)
    if ((Get-Settings).pet) { [void](Start-Pet) }

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 2000
    $timer.add_Tick({
            try {
                Update-TrayState
                Start-WatcherIfNeeded
                Sync-PetWithSettings
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

function Stop-PetProcess {
    # 只停「单独跑的桌宠进程」(-Mode pet);托盘自带的那个跟着托盘一起停
    foreach ($p in (Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue)) {
        if ($p.ProcessId -ne $PID -and $p.CommandLine -match 'Connect-CampusNet\.ps1' -and $p.CommandLine -match '\-Mode\s+pet') {
            try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop; Write-Log ('已停止独立桌宠 pid=' + $p.ProcessId) 'OK' } catch { }
        }
    }
}

function Start-PetStandalone {
    # 「只保留桌宠」:托盘和守护都停掉,单开一个隐藏的桌宠进程,关掉窗口也不会把它带走
    foreach ($p in (Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue)) {
        if ($p.ProcessId -ne $PID -and $p.CommandLine -match 'Connect-CampusNet\.ps1' -and $p.CommandLine -match '\-Mode\s+pet') { return $true }
    }
    try {
        $s = Get-Settings
        if (-not [bool]$s.pet) {
            $s.pet = $true
            Save-Settings -Settings $s
        }
        Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Mode pet -Background' -f $SelfScript)
        Write-Log '已单独拉起桌宠进程(托盘与守护已停)' 'OK'
        return $true
    } catch {
        Write-Log ('单独拉起桌宠失败:' + $_.Exception.Message) 'WARN'
        return $false
    }
}

function Show-ExitChooser {
    # 关程序窗口前问一句:完全退出 / 只留桌宠 / 最小化
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '退出校园网自动登录'
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.TopMost = $true
    $f.ClientSize = [System.Drawing.Size]::new(472, 286)

    $lb = New-Object System.Windows.Forms.Label
    $lb.Text = '程序窗口要关掉了,你想留下什么?'
    $lb.Location = [System.Drawing.Point]::new(18, 14)
    $lb.Size = [System.Drawing.Size]::new(436, 26)
    try { $lb.Font = [System.Drawing.Font]::new('Microsoft YaHei UI', 10.5, [System.Drawing.FontStyle]::Bold) } catch { }
    $f.Controls.Add($lb)

    $script:ExitChoice = 'cancel'
    foreach ($o in @(
            @{ t = '完全退出程序 —— 托盘、守护、桌宠全停'; v = 'quit'; y = 54 },
            @{ t = '只保留桌宠 —— 停托盘和守护,桌宠留在桌面'; v = 'pet'; y = 106 },
            @{ t = '最小化程序 —— 窗口收起来,一切照常跑'; v = 'min'; y = 158 }
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
    $bCancel.Location = [System.Drawing.Point]::new(18, 232)
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
            Stop-PetProcess
            Write-Stdout ''
            Write-Stdout '已完全退出:守护、托盘、桌宠都停了。'
            Write-Stdout '开机自启还在,下次登录 Windows 会自己起来(不想要就选 10 卸载开机自启)。'
            Write-Log '用户选择:完全退出程序' 'OK'
            return 'exit'
        }
        'pet' {
            Stop-Watcher
            [void](Stop-Tray)
            Start-Sleep -Milliseconds 400
            Write-Stdout ''
            if (Start-PetStandalone) {
                Write-Stdout '只留下桌宠:托盘和守护已停,桌宠继续待命(右键桌宠可以打开菜单 / 隐藏它)'
            } else {
                Write-Stdout '想只留桌宠,但单独拉起桌宠失败 —— 看日志(菜单 L)里的原因'
            }
            Write-Log '用户选择:只保留桌宠' 'OK'
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
        Write-Stdout '                校园网自动登录'
        Write-Stdout ('                      v' + $script:AppVersion)
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
        Write-Stdout ('  桌面宠物: ' + $(if ((Get-Settings).pet) { '开' } else { '关' }) + '   素材: ' + $script:PetUrl)
        $wxLine = Format-PetWeather
        if ($wxLine) {
            Write-Stdout ('  天气: ' + $wxLine)
        } else {
            Write-Stdout '  天气: (还没有数据,按 W 填个城市)'
        }
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
        Write-Stdout '  P) 桌面宠物: 开/关切换'
        Write-Stdout '  W) 设置天气城市(留空=按 IP 定位)'
        Write-Stdout '  O) 门户类型 / 通用认证配置'
        $curProto = [string](Get-PortalProfile).name
        Write-Stdout ('       (当前: ' + $curProto + ')')
        Write-Stdout '  C) 采集诊断包(远程排查用,自动复制到剪贴板)'
        Write-Stdout '  G) 新手向导(问一句是不是校园网 + 认门户)'
        Write-Stdout '  L) 查看日志'
        Write-Stdout '  0) 退出(可选:完全退出 / 只保留桌宠 / 最小化)'
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
            'P' {
                $on = -not [bool](Get-Settings).pet
                Set-PetEnabled -On $on
                Write-Stdout ('桌面宠物: ' + $(if ($on) { '开(托盘那边 2 秒内会出来;素材来自 ' + $script:PetUrl + ')' } else { '关' }))
            }
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
            'W' {
                $curCity = [string](Get-PetSetting -Name 'petWeatherCity' -Default '')
                $curToken = [string](Get-PetSetting -Name 'caiyunToken' -Default '')
                Write-Stdout ('当前天气城市: ' + $(if ($curCity) { $curCity } else { '(按 IP 自动定位)' }))
                Write-Stdout ('彩云天气 token: ' + $(if ($curToken) { '已配置(' + $curToken.Substring(0, [Math]::Min(4, $curToken.Length)) + '…)' } else { '未配置(用 Open-Meteo)' }))
                $nc = Read-Host '输入城市名/区县名(回车取消,输入 - 表示改回按 IP 定位)'
                if (-not [string]::IsNullOrWhiteSpace($nc)) {
                    $nc = $nc.Trim()
                    if ($nc -eq '-') { $nc = '' }
                    Set-PetSettingValue -Name 'petWeatherCity' -Value $nc
                    $script:PetWeather = $null
                    $script:PetWeatherAt = [datetime]::MinValue
                    $w2 = Update-PetWeather
                    if ($w2.ok) {
                        Write-Stdout ('天气: ' + $w2.text + ' ' + $w2.temp + '°C · ' + $w2.place + ' · ' + $w2.dressing + '  [' + $w2.source + ']')
                    } else { Write-Stdout '没查到(城市名写对了吗?或者网络被拦)' }
                    Write-Log ('天气位置设为: ' + $(if ($nc) { $nc } else { '(按 IP 定位)' }))
                }

                $nt = Read-Host '彩云天气 token(回车保持;输入 - 表示清空)'
                if (-not [string]::IsNullOrWhiteSpace($nt)) {
                    $nt = $nt.Trim()
                    if ($nt -eq '-') { $nt = '' }
                    Set-PetSettingValue -Name 'caiyunToken' -Value $nt
                    Write-Log ('彩云 token ' + $(if ($nt) { '已更新' } else { '已清空' }))
                    $script:PetWeather = $null
                    $script:PetWeatherAt = [datetime]::MinValue
                    $w3 = Update-PetWeather
                    if ($w3.ok) { Write-Stdout ('天气: ' + $w3.text + ' ' + $w3.temp + '°C · ' + $w3.place + ' · ' + $w3.dressing + '  [' + $w3.source + ']') }
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

    'weather' {
        # 联网识别位置 + 取当前天气,结果写 pet-weather.json 给悬停面板用
        if ($CaiyunToken) { Set-PetSettingValue -Name 'caiyunToken' -Value $CaiyunToken }
        $w = Update-PetWeather
        if ($w.ok) {
            Write-Stdout ('天气: ' + $w.text + ' ' + $w.temp + '°C(体感 ' + $w.feels + '°C) · ' + $w.place)
            Write-Stdout ('穿衣: ' + $w.dressing)
            Write-Stdout ('来源: ' + $w.source + $(if ($w.alert) { '   预警: ' + $w.alert } else { '' }))
            exit 0
        }
        Write-Stdout '拿不到位置/天气(可能没网或被拦)'
        exit 1
    }

    'pet' {
        # 单独把宠物跑起来(调试/预览用);平时它是跟着托盘走的
        if ($PetOff) {
            $s = Get-Settings
            $s.pet = $false
            Save-Settings -Settings $s
            Write-Stdout '桌面宠物已关闭(托盘那边 2 秒内会收掉)'
            exit 0
        }
        Add-Type -AssemblyName System.Windows.Forms
        if (-not (Start-Pet)) { exit 1 }
        $s = Get-Settings
        $s.pet = $true
        Save-Settings -Settings $s
        Write-Stdout ''
        Write-Stdout ('宠物素材来自 dsh-pet(MIT):' + $script:PetUrl)
        Write-Stdout '右键小宠物有菜单;关掉这个窗口(Ctrl+C)就退出'
        [System.Windows.Forms.Application]::Run()
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
        # 选过「完全退出 / 只保留桌宠」就别再把托盘拉回来了
        if ($script:MenuExit -eq 'exit') { exit 0 }
        # 菜单关掉之后:装了自启就把托盘拉回来;没装就明确告诉用户为啥没有常驻图标
        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            [void](Start-TrayIfNeeded)
        } else {
            Write-Stdout ''
            Write-Stdout '提示: 还没装「开机自启」,所以关掉这个窗口后不会有托盘图标和桌宠常驻。'
            Write-Stdout '      菜单里选 9) 安装开机自启 就会一直留在右下角。'
        }
    }
}
