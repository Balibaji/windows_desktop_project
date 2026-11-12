# AccessClient + mstsc 的 `127.0.0.1` 双端口隧道：原理、验证与优雅交互

> 适用场景：通过企业堡垒机（如齐治 / H3C SecPath A2000 / AK 系列运维审计系统）访问内网 Windows 资产，
> 本地看到 `AccessClient` 拉起 `mstsc`，日志里出现两个 `127.0.0.1:<port>` 监听端口。

---

## 0. 结论速览（TL;DR）

1. **`AccessClient` 不是远程桌面客户端**，而是堡垒机的本地组件：负责“拉起本地客户端 + 建立隧道 + 代填凭据 + 上报审计”。
2. **`mstsc` 连接的不是目标服务器，而是本机的 `127.0.0.1:<数据端口>`**。真正的目标地址、账号、密码都藏在隧道另一端。
3. **两个 `127.0.0.1` 端口通常是“数据面 + 控制面”**：
   - 数据面：`mstsc` 连的那个 TCP 代理端口；
   - 控制面：浏览器 / 本地 IPC / 心跳 / 审计上报用的本地端口。
   - （少数情况是“两条独立转发”，或 IPv4/IPv6 双栈各绑一个。）
4. **优雅交互的核心**：把 loopback 端口当成**服务端点**来用，而不是去点 GUI。
   - 先 `Get-NetTCPConnection` 定位端口与属主进程；
   - 再 `mstsc /v:127.0.0.1:<port>` 或直接复用已打开的会话；
   - 自动化前先**等端口就绪**，不要 `Start-Sleep` 赌时间；
   - 能走官方 CLI / 直连 / WinRM 就不要用 SendInput + OCR 点画面。

---

## 1. 典型拓扑与数据流

```
┌─────────────┐   自定义协议 / 本地 HTTP     ┌──────────────┐
│  浏览器      │ ───────────────────────────> │ AccessClient │
│ (堡垒机 Web) │ <─────────────────────────── │  (本地 agent)│
└─────────────┘                              └──────┬───────┘
                                                    │ 出站长连接
                                                    │ (TLS / SSH / 私有协议)
                                                    ▼
                                             ┌──────────────┐
                                             │   堡垒机      │
                                             │ (运维审计系统) │
                                             └──────┬───────┘
                                                    │ RDP 3389
                                                    ▼
                                             ┌──────────────┐
                                             │ 目标 Windows  │
                                             └──────────────┘

本机内部：
  AccessClient  监听 127.0.0.1:P_ctrl   ← 控制 / IPC / 审计 / 心跳
  AccessClient  监听 127.0.0.1:P_data   ← TCP 代理（真正的数据面）
                       ▲
                       │ TCP
                  mstsc /v:127.0.0.1:P_data
```

数据流（点击“访问资产”之后）：

1. 浏览器把“要访问哪个资产、用哪个账号、什么分辨率/剪贴板策略”交给 `AccessClient`（自定义 URL 协议或本地 HTTP 回调）。
2. `AccessClient` 与堡垒机建立一条**出站长连接**（只出不进，天然穿防火墙）。
3. `AccessClient` 在本机 `127.0.0.1` 起监听端口，作为这条会话的**本地端点**。
4. 它用 `mstsc /v:127.0.0.1:<port>`（或临时 `.rdp` 文件）拉起远程桌面。
5. `mstsc` 连本地端口 → `AccessClient` 在长连接里开一条流 → 堡垒机 → 目标机 `3389`。
6. 之后键盘、鼠标、剪贴板、磁盘重定向全部走这条流，堡垒机顺带做审计/录像/水印。

---

## 2. `AccessClient` 到底是什么

以齐治堡垒机 / H3C SecPath A2000（AK 系列）为例：

| 项目 | 说明 |
| --- | --- |
| 定位 | 运维审计系统的**本地客户端组件**，装在被运维人员的 PC 上 |
| 获取方式 | 登录 Web → “访问资产”提示下载；或“帮助 → AccessClient → 下载”手工安装 |
| 主要职责 | ① 按协议拉起本地客户端（mstsc / SecureCRT / Xshell / FileZilla…）<br>② 建立并维持到堡垒机的会话隧道<br>③ 代填账号密码 / OTP<br>④ 下发剪贴板、磁盘映射、分辨率等策略<br>⑤ 采集审计数据 |
| 跨平台 | Windows / macOS / 麒麟 / 统信；Linux 下自带 `freerdp`，Windows 下调用系统 `mstsc` |
| 日志位置 | `%TEMP%\AccessClient.log`（例如 `C:\Users\<用户>\AppData\Local\Temp\AccessClient.log`） |
| 注册表线索 | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\<客户端>.exe`（记录本地客户端路径） |

它**不是** UI 自动化工具，也不是 VPN。它更像一个“**按需启动的本地隧道代理 + 客户端启动器**”。

同一套东西在 H3C 文档里对应两种访问模式：

- **RDP 直连**：`mstsc` 先连堡垒机（如 `堡垒机IP:53389`），在堡垒机界面里再选资产；
- **RDP 穿透**：`mstsc` 直接连目标，用户名写成 `堡垒机用户/资产IP/资产账号`，由堡垒机转发。

而 `AccessClient` 模式下，这两步被折叠成一个动作：**由 agent 起本地端口，`mstsc` 连 `127.0.0.1`**。

---

## 3. 两个 `127.0.0.1` 端口分别是什么

先明确一件事：**`127.0.0.1` 是“本机回环”，不是远端。** 所以这两个端口一定是你本机上的某个进程（大概率就是 `AccessClient`）在监听。

### 3.1 三种常见分工

| 假设 | 端口 1 | 端口 2 | 如何验证 |
| --- | --- | --- | --- |
| **A. 数据 + 控制**（最常见） | TCP 代理，`mstsc` 连它 | 浏览器/IPC/心跳/审计 | 看谁在 `ESTABLISHED` 连 P1；P2 是否只有本机短连接 |
| **B. 两条独立转发** | RDP 转发 | 剪贴板/文件/辅助通道 | 两个端口都有稳定的 `ESTABLISHED` 到堡垒机 |
| **C. 双栈绑定** | `127.0.0.1:P` | `::1:P`（同一端口号） | 端口号相同、地址不同，直接忽略其一 |

### 3.2 为什么一定要绑 `127.0.0.1`

- **最小暴露面**：只有本机能连，同网段的人蹭不到你的会话。
- **免防火墙规则、免管理员权限**：用 >1024 端口即可，不需要 `netsh` 或管理员。
- **一会话一端点**：端口随机化，天然隔离多会话，也方便按会话审计。
- **与零信任模型一致**：本地只是“会话端点”，真正的鉴权在网关。

---

## 4. 实现机制拆解

### 4.1 本地监听 + TCP 代理

核心就是一个 `socket.listen()` + `accept()` 循环：

```text
listen(127.0.0.1, P)
  └─ accept() 得到 local_sock
       └─ 在已有的出站长连接上 open_stream()
            └─ 告诉堡垒机：connect(target_ip, 3389)
                 └─ 双向 pipe(local_sock, stream)
```

所以从 `mstsc` 的视角看，它只是连了一个普通的 RDP 服务器；从堡垒机视角看，它只是在服务一条被授权的隧道流。

### 4.2 传输层可能是什么

按实现成熟度和企业习惯，常见三种：

1. **复用 SSH**：AccessClient 内部起一个 SSH 客户端连堡垒机，用 `direct-tcpip` channel 转发。
   - 旁证：堡垒机通常会在 `sshd` 侧限制 channel 类型，自己手搓 `ssh -L` 可能报
     `channel 0: open failed: administratively prohibited`。
2. **TLS 长连接 + 多路复用**（yamux / smux / HTTP2 CONNECT）：一条连接承载多路流，控制帧与数据帧分离——这也解释了为什么会看到两个端口。
3. **私有协议**：厂商自研，抓包只能看到 TLS。

### 4.3 `mstsc` 是怎么被拉起来的

两种方式，都会在进程命令行或临时文件里留下痕迹：

```text
mstsc.exe /v:127.0.0.1:52341
```

或者生成一个临时 `.rdp` 文件（`full address:s:127.0.0.1:52341`），里面同时带上：

- `screen mode id` / `desktopwidth` / `desktopheight`（堡垒机下发的分辨率）
- `drivestoredirect:s:`（磁盘映射策略）
- `redirectclipboard:i:1`（剪贴板策略）
- `authentication level` / `prompt for credentials`
- 有时还有 `password 51:b:<DPAPI 密文>`（凭据代填）

### 4.4 凭据是怎么“免密”的

按可靠程度排序：

| 方式 | 特征 |
| --- | --- |
| 服务端代填 | AccessClient 只在协议层转发，堡垒机侧用托管账号登录目标机 |
| `cmdkey` / 凭据管理器 | `cmdkey /list` 能看到目标条目 |
| `.rdp` 里的 `password 51:b:` | DPAPI 加密，只有当前用户能解 |
| 自定义 Credential Provider / 注入 | `mstsc` 不弹登录框，进程内有额外模块 |

> 调试时可以看 `.rdp` 内容确认“到底连的是哪个本地端口”，但**不要**去解密或搬运凭据——那已经越过审计边界了。

### 4.5 为什么不让 `mstsc` 直连目标

- 隐藏真实资产地址与端口；
- 统一鉴权（SSO / OTP / 动态授权）；
- 强制审计（键盘记录、会话录像、水印）；
- 策略下发（剪贴板、磁盘、分辨率、超时）；
- 协议适配（Linux 下换成自带 `freerdp`）。

---

## 5. 现场验证清单（照着敲）

### 5.1 找到端口与属主进程

```powershell
Get-NetTCPConnection -State Listen |
  Where-Object LocalAddress -in '127.0.0.1','::1' |
  Select-Object LocalAddress,LocalPort,OwningProcess,
    @{n='Proc';e={(Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue).ProcessName}} |
  Sort-Object LocalPort | Format-Table -AutoSize
```

### 5.2 看进程命令行与父进程

```powershell
# AccessClient 本体
Get-CimInstance Win32_Process |
  Where-Object { $_.Name -match 'AccessClient' } |
  Select-Object ProcessId,ParentProcessId,Name,ExecutablePath,CommandLine | Format-List

# mstsc 的启动参数（关键：/v:127.0.0.1:端口）
Get-CimInstance Win32_Process -Filter "Name='mstsc.exe'" |
  Select-Object ProcessId,ParentProcessId,CommandLine | Format-List

# 父进程链：确认 mstsc 是 AccessClient 拉起的
$p = Get-CimInstance Win32_Process -Filter "Name='mstsc.exe'" | Select-Object -First 1
while ($p) {
  "{0}  {1}  {2}" -f $p.ProcessId, $p.Name, $p.CommandLine
  $p = Get-CimInstance Win32_Process -Filter "ProcessId=$($p.ParentProcessId)" -ErrorAction SilentlyContinue
}
```

### 5.3 看端口上到底连着谁

```powershell
# 谁在连这个数据端口（应该是 mstsc）
Get-NetTCPConnection -LocalPort <P_data> |
  Select-Object State,LocalAddress,LocalPort,RemoteAddress,RemotePort,OwningProcess,
    @{n='Proc';e={(Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue).ProcessName}}

# 端口探活 + 指纹
Test-NetConnection 127.0.0.1 -Port <P_data> -InformationLevel Detailed
```

### 5.4 看临时 `.rdp` 文件

```powershell
Get-ChildItem $env:TEMP, "$env:USERPROFILE\Downloads" -Filter *.rdp -Recurse -ErrorAction SilentlyContinue |
  Sort-Object LastWriteTime -Descending | Select-Object -First 5 FullName,LastWriteTime
```

打开后重点看：

```ini
full address:s:127.0.0.1:52341
screen mode id:i:1
desktopwidth:i:1920
desktopheight:i:1080
drivestoredirect:s:
redirectclipboard:i:1
```

### 5.5 看日志

```powershell
Get-Content "$env:TEMP\AccessClient.log" -Tail 100 -Wait
```

日志里通常能看到：端口分配、目标地址、协议、隧道建立/断开、代填结果、错误码。

### 5.6 判断表

| 现象 | 结论 |
| --- | --- |
| `mstsc` 命令行是 `/v:127.0.0.1:xxxxx` | 确认是“本地端口 + agent 隧道”模式 |
| `P_data` 上有 `mstsc.exe` 的 `ESTABLISHED` | 数据面端口找对了 |
| `P_ctrl` 只有本机短连接 / 无外部连接 | 控制面 / IPC，别去连它当 RDP |
| 端口号每次启动都变 | 会话式端点，**禁止写死** |
| 两个端口号相同、地址分别是 `127.0.0.1` 与 `::1` | 双栈绑定，同一个服务 |
| `AccessClient.log` 里出现 `listen 127.0.0.1` | 直接印证本地监听 |

---

## 6. 怎么优雅地交互

核心原则：**把 loopback 端口当成一个服务端点，用协议去说话，而不是用鼠标去点。**

### 6.1 先选路：能不走 RDP 就不走 RDP

```
需要跑命令/脚本？
  ├─ 能 SSH 或 WinRM  → 走命令行（最稳）
  ├─ 只能 RDP 但只需要执行命令 → 隧道 + WinRM / 计划任务 / PsExec
  └─ 只有图形界面 → 才用 mstsc + 键鼠/截图

需要长期无人值守？
  └─ 在目标机/跳板机装 agent，本机通过 HTTP/socket 下指令，不要在本机截 mstsc
```

### 6.2 复用已有会话，而不是再造一个

`AccessClient` 已经帮你建好隧道时，**不要**自己去 `ssh -L` 或另起客户端：

```powershell
# 直接用已打开的 mstsc 窗口
python rdp_ops.py status     --any --process mstsc
python rdp_ops.py screenshot --any --process mstsc --out shot.png
python rdp_ops.py click-text --any --process mstsc --text "确定"
```

原因：端口通常**一次会话一个**，而且与审计记录绑定；自己抢连可能被拒绝，或在审计侧留下异常。

### 6.3 需要自己起会话时：用官方入口

```powershell
# 如果堡垒机支持 RDP 直连/穿透，优先用它文档里的标准端口
mstsc /v:<堡垒机IP>:53389
mstsc /v:<目标IP>            # 用户名按厂商格式，例如 H3C：opt/10.2.234.21/administrator
```

若确认底层是 SSH 且堡垒机**允许** `direct-tcpip`，才考虑等价自建：

```bash
ssh -N -L 127.0.0.1:13389:10.0.0.8:3389 user@bastion
```

并用配置代替手写参数：

```ssh-config
Host bastion
    HostName bastion.example.com
    User yourname
    ServerAliveInterval 20
    ServerAliveCountMax 3
    ExitOnForwardFailure yes
    ControlMaster auto
    ControlPath ~/.ssh/cm-%r@%h:%p
    ControlPersist 10m

Host target
    HostName 10.0.0.8
    User admin
    ProxyJump bastion
```

> 如果报 `administratively prohibited`，说明堡垒机禁用了转发通道，**只能**走 AccessClient。

### 6.4 端口发现要“等”，不要“睡”

```powershell
function Wait-LoopbackPort {
    param([int]$Port, [int]$TimeoutSec = 30)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        if (Test-NetConnection 127.0.0.1 -Port $Port -InformationLevel Quiet) { return $true }
        Start-Sleep -Milliseconds 300
    }
    return $false
}
```

比 `Start-Sleep 5` 稳得多，而且失败时能给出明确错误。

### 6.5 自动化脚本里的正确姿势

```powershell
# 1) 从 mstsc 命令行解析出实际端口，而不是猜
$cmd = (Get-CimInstance Win32_Process -Filter "Name='mstsc.exe'" |
        Select-Object -First 1).CommandLine
$port = if ($cmd -match '/v:127\.0\.0\.1:(\d+)') { [int]$Matches[1] } else { $null }

# 2) 等端口就绪
if (-not $port -or -not (Wait-LoopbackPort -Port $port -TimeoutSec 20)) {
    throw "RDP 本地端口未就绪（mstsc 未连上或 AccessClient 隧道未建立）"
}

# 3) 再交给自动化工具
python rdp_ops.py find --any --process mstsc --json
```

### 6.6 已实现：`tunnel` / `wait-port` / `health` 子命令

`rdp_ops.py` 与 `RdpOps.ps1` 现在都内置了隧道感知，不需要再自己拼 `Get-NetTCPConnection`：

```bash
# 发现隧道：端口、属主进程、父进程、是否已连接
python rdp_ops.py tunnel
python rdp_ops.py tunnel --app AccessClient
python rdp_ops.py tunnel --watch --json          # 持续监控，状态变化才输出

# 等端口就绪（AccessClient 建隧道需要时间）
python rdp_ops.py wait-port --port 52341 --timeout 20

# 综合健康检查：窗口存在 / 未挂起 / 未最小化 / 端口 LISTEN / mstsc 已连接
python rdp_ops.py health --any --process mstsc
python rdp_ops.py health --any --process mstsc --screen --json
```

```powershell
.\RdpOps.ps1 tunnel
.\RdpOps.ps1 tunnel -App "AccessClient*" -Watch
.\RdpOps.ps1 wait-port -Port 52341 -TimeoutSec 20
.\RdpOps.ps1 health -Any -ProcessName mstsc -Screen -Json
```

实现要点：

- `tunnel` 同时解析 **命令行 `/v:`** 和 **该进程到回环地址的 `ESTABLISHED` 连接**，所以用 `.rdp` 文件启动也能识别端口；
- `wait-port` 只查 `LISTEN`，默认**不**主动连 RDP 端口（避免产生额外握手/审计记录），需要时加 `--connect` / `-Connect`；
- `health` 返回非零退出码，可直接用于流水线判断；`--screen` 用连续两张截图的 SHA256 判断画面是否卡死；
- 端口属主、父进程名会一并输出，便于确认“这条隧道确实是 AccessClient 建的”。

### 6.7 健康检查与重连

| 检查项 | 方法 | 失败含义 |
| --- | --- | --- |
| 隧道端口还在 | `Test-NetConnection 127.0.0.1 -Port <P>` | AccessClient 退出 / 会话过期 |
| TCP 仍连着 | `Get-NetTCPConnection -LocalPort <P> -State Established` | 长连接断了（mstsc 画面会冻结） |
| 窗口还在 | `Get-Process mstsc` + 句柄有效 | 用户关掉了窗口 |
| 画面在刷新 | 连续两次截图 hash 不同 | 会话卡死/黑屏 |

> 断线后不要盲目重试同一端口——端口可能已被回收，正确做法是让 `AccessClient` 重新发起会话。

### 6.8 安全与合规边界

- **绝不**把监听改成 `0.0.0.0`，那等于把隧道暴露给整个局域网。
- Windows 的 loopback 是**全机共享**的，多用户机器上其他人理论上也能连你的 `127.0.0.1:P`。
- 不要试图 dump / 转发 `AccessClient` 的凭据、token、DPAPI 密文。
- 堡垒机通常有键盘记录、会话录像、屏幕水印、会话时长上限；**自动化操作同样会被审计**，保持行为可解释。
- 只用厂商支持的客户端与版本，避免“审计记录不全”导致合规问题。

---

## 7. 常见坑

| 坑 | 表现 | 处理 |
| --- | --- | --- |
| 端口写死 | 换了会话就连不上 | 每次从进程命令行/日志解析 |
| 端口只在当前用户会话有效 | 切换用户/锁屏后失效 | 不要跨会话复用 |
| `AccessClient` 升级/退出 | 端口消失，`mstsc` 冻结 | 监听进程退出事件，重新发起会话 |
| 一个端口只服务一个会话 | 第二个 `mstsc` 连不上 | 每个会话用各自的端口 |
| 分辨率/智能缩放 | 截图坐标漂移 | 固定窗口尺寸与远端分辨率，关掉智能缩放 |
| DPI 缩放 | 点击偏位 | 脚本声明 DPI aware，坐标统一按客户区换算 |
| 剪贴板/磁盘重定向被策略禁用 | 粘贴无效、盘符看不到 | 找管理员开权限，别硬绕 |
| 本机锁屏 | `SendInput` 失效 | 自动化期间保持解锁 |
| 目标会话锁屏/断开 | 画面黑、输入无效 | 远端 `tscon ... /dest:console` 或保持会话活跃 |
| 多显示器/全屏 | 客户区坐标与屏幕坐标混淆 | 统一用 `ClientToScreen` 换算 |
| IPv6 双栈 | 看到两个“相同端口” | 只取 `127.0.0.1` 那条 |
| 杀软拦截 | `AccessClient` 起不来 | 加白名单，用官方安装包 |

---

## 8. 和本项目脚本的对接

| 目的 | 命令 |
| --- | --- |
| 发现隧道端口 | `python rdp_ops.py tunnel --app AccessClient` |
| 等端口就绪 | `python rdp_ops.py wait-port --port 52341 --timeout 20` |
| 综合健康检查 | `python rdp_ops.py health --any --process mstsc --screen` |
| 确认窗口在不在 | `python rdp_ops.py status --any --process mstsc` |
| 截客户区 | `python rdp_ops.py screenshot --any --process mstsc --out shot.png` |
| OCR 识别 | `python rdp_ops.py ocr --any --process mstsc --language zh-CN` |
| 点文字 | `python rdp_ops.py click-text --any --process mstsc --text "确定"` |
| 等文字出现 | `python rdp_ops.py wait-text --any --process mstsc --text "完成" --timeout 30` |
| 发按键 | `python rdp_ops.py keys --any --process mstsc --keys ctrl+v enter` |
| 剪贴板 | `python rdp_ops.py clip-set --text "..."` → `paste` |

PowerShell 版同理：

```powershell
.\RdpOps.ps1 tunnel     -App "AccessClient*"
.\RdpOps.ps1 wait-port  -Port 52341 -TimeoutSec 20
.\RdpOps.ps1 health     -Any -ProcessName mstsc -Screen -Json
.\RdpOps.ps1 status     -Any -ProcessName mstsc -Json
.\RdpOps.ps1 screenshot -Any -ProcessName mstsc -Out shot.png
.\RdpOps.ps1 click-text -Any -ProcessName mstsc -Text "确定" -DryRun
```

注意：`rdp_ops.py` 的 `--any` 只是放开“必须是 RDP 进程”的默认过滤，**不会**帮你建立隧道。
隧道由 `AccessClient` 负责；脚本只负责“找到窗口 + 输入 + 截图 + OCR”。

---

## 9. 附录：速查

```powershell
# 端口 → 进程
Get-NetTCPConnection -State Listen |
  ? LocalAddress -in '127.0.0.1','::1' |
  select LocalPort,OwningProcess,@{n='P';e={(Get-Process -Id $_.OwningProcess).ProcessName}}

# 进程 → 命令行
Get-CimInstance Win32_Process -Filter "Name='mstsc.exe'" | select ProcessId,ParentProcessId,CommandLine

# 端口 → 连接
Get-NetTCPConnection -LocalPort <P> | select State,RemoteAddress,RemotePort,OwningProcess

# 探活
Test-NetConnection 127.0.0.1 -Port <P> -InformationLevel Detailed

# 日志
Get-Content "$env:TEMP\AccessClient.log" -Tail 100 -Wait

# 直连
mstsc /v:127.0.0.1:<P>
```

```bash
# 等价 SSH 隧道（仅在堡垒机允许 direct-tcpip 时）
ssh -N -L 127.0.0.1:13389:10.0.0.8:3389 user@bastion
ssh -J user@bastion user@10.0.0.8
```

---

## 10. 一句话总结

**两个 `127.0.0.1` 端口 = “本机触发的出站隧道”的本地端点**（`AccessClient` 建立，`mstsc` 使用）。
实现上是“本地 listen + 出站长连接开流 + 堡垒机侧 connect 目标”。
交互上把它当 API 端点：用官方入口、等端口就绪、复用已建会话、能走命令行就别碰 RDP 位图。
