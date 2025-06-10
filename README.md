# Windows 10/11自动关机工具

一个自动监控网络连接状态的Windows工具，在网络断开且用户无活动一段时间后自动关机。

## 功能特点

- 实时监控网络连接状态
- 检测用户键盘和鼠标活动
- 网络断开且无用户活动一段时间后自动关机
- 支持设置空闲时间阈值
- 支持创建Windows计划任务，开机自启动

## 使用方法

### 基本用法

直接运行程序，默认设置为网络断开后2分钟无活动将自动关机：

```
LiuMengAutoPowerOffAfterOffline.exe
```

### 命令行参数

- `-i, --idle`: 设置空闲时间阈值（分钟），默认为2分钟
- `-t, --task`: 创建Windows计划任务，系统启动时自动运行

### 示例

设置空闲时间为5分钟：

```
LiuMengAutoPowerOffAfterOffline.exe -i 5
```

创建Windows计划任务：

```
LiuMengAutoPowerOffAfterOffline.exe -t
```

同时设置空闲时间并创建任务：

```
LiuMengAutoPowerOffAfterOffline.exe -i 5 -t
```

## 工作原理

1. 程序启动后，会监控网络连接状态
2. 当检测到网络断开时，开始计时
3. 如果在设定的时间内没有检测到用户活动（键盘或鼠标），则执行关机命令
4. 如果在此期间网络恢复连接或检测到用户活动，则取消关机计划

## 系统要求

- Windows操作系统
- 管理员权限（创建计划任务时需要）

## 授权协议

本仓库采用的是Apache License 2.0协议，详情请看[LICENSE](LICENSE)文件。