Windows Internet Sharing Enabler
================================

使用 Powershell 启用 Windows 7+ 自带的热点功能。需要支持此功能的无线网卡。

## 使用方法

所有方法需要管理员权限。

### 自定义配置

下载 [ps/enablesharing.ps1](ps/enablesharing.ps1)，修改前三个变量：

```shell
$name = "已经连接到互联网的适配器名，无线接入的设备将使用这一网络接入互联网"
$ssid = "无线网络 SSID"
$key = "无线网络密码"
```

其中 `$name` 可以通过 `netsh interface show interface` 命令获取。

-------------------------------

以下方法不推荐，仅供参考。会使用这些默认设置：
```shell
$name = "以太网"
$ssid = "testssid"
$key = "password"
```

### 直接从 cmd 运行

```shell
@powershell -NoProfile -ExecutionPolicy unrestricted -Command "iex ((new-object net.webclient).DownloadString('https://raw.githubusercontent.com/Jamesits/windows-internet-sharing-enabler/master/ps/enablesharing.ps1'))"
```

### 直接从 Powershell 运行

```shell
iex ((new-object net.webclient).DownloadString('https://raw.githubusercontent.com/Jamesits/windows-internet-sharing-enabler/master/ps/enablesharing.ps1'))
```

## 查询热点运行状态

`netsh wlan show interfaces` 可以查看当前 Wi-Fi Adapter 属性。

`netsh wlan show hostednetwork` 可以查看已连接的设备列表。

## 常见故障处理

```
必须使用管理员权限从命令提示符处运行此命令。
```

没给管理员权限。

-------------------------------

```
无法启动承载网络。
组或资源的状态不是执行请求操作的正确状态。
```

请运行一下 `netsh wlan show drivers`。

显示 `系统上没有无线接口。`的话，说明你的机子不支持 Wi-Fi。

否则请检查是否输出了 `支持的承载网络： 是`，如果没有的话说明你的无线网卡不支持这一功能。

-------------------------------

```
使用“1”个参数调用“Invoke”时发生异常:“无法处理参数，因为参数“arguments”的值为空。请将参数“arguments”的
非空值。”
所在位置 enablesharing.ps1:35 字符: 1
+ $config = $m.INetSharingConfigurationForINetConnection.Invoke($c)
+ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
+ CategoryInfo          : NotSpecified: (:) [], MethodInvocationException
+ FullyQualifiedErrorId : PSArgumentNullException

不能对 Null 值表达式调用方法。
所在位置 enablesharing.ps1:46 字符: 1
+ $config.EnableSharing(0)
+ ~~~~~~~~~~~~~~~~~~~~~~~~
+ CategoryInfo          : InvalidOperation: (:) []，RuntimeException
+ FullyQualifiedErrorId : InvokeMethodOnNull
```

一般都是 `$name` 填错了，或者没给管理员权限。

## 关于

本来打算做个 VB 6.0 程序当壳的，netsh 这东西的输出不是 Machine readable 的，Powershell 与程序的交互也极麻烦（尤其是 VB 6.0 没有内建 pipe 支持）还有兼容性问题，遂放弃。完善了下脚本单独拿出来用。

<!-- Written for ikeltis -->

by James Swineson. 2014-12-13
