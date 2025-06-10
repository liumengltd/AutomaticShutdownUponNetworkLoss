package main

import (
	"fmt"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
	"syscall"
	"time"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/spf13/pflag"
)

// Configuration
const (
	ShutdownCmd     = "shutdown /s /t 0" // 关机命令
	IdleTimeMinutes = 2                  // 关机前侦听输入时长（分钟）
)

// Windows virtual key codes
const (
	VkLButton = 0x01
	VKRButton = 0x02
	VKMButton = 0x04
)

// KeyPressed Windows API constants
const (
	KeyPressed = 0x8000
)

// GetAsyncKeyState calls the Windows API function
func GetAsyncKeyState(vKey int) int {
	ret, _, _ := syscall.SyscallN(
		getProcAddress("user32.dll", "GetAsyncKeyState"),
		uintptr(vKey),
		0,
		0,
	)
	return int(ret)
}

// getProcAddress 获取进程地址
func getProcAddress(dll, proc string) uintptr {
	handle := syscall.MustLoadDLL(dll)
	return handle.MustFindProc(proc).Addr()
}

// ActivityMonitor 鼠标和键盘活动侦听器
type ActivityMonitor struct {
	lastActivityTime time.Time
	running          bool
}

// NewActivityMonitor 创建一个活动侦听器
func NewActivityMonitor() *ActivityMonitor {
	return &ActivityMonitor{
		lastActivityTime: time.Now(),
		running:          true,
	}
}

// Start 启动活动侦听器
func (am *ActivityMonitor) Start() {
	go am.keyboardListener()
	go am.mouseListener()
}

// keyboardListener 键盘侦听器
func (am *ActivityMonitor) keyboardListener() {
	for am.running {
		for key := 8; key < 256; key++ {
			if (GetAsyncKeyState(key) & KeyPressed) != 0 {
				am.lastActivityTime = time.Now()
				break
			}
		}
		time.Sleep(100 * time.Millisecond)
	}
}

// mouseListener 鼠标侦听器
func (am *ActivityMonitor) mouseListener() {
	for am.running {
		if (GetAsyncKeyState(VkLButton)&KeyPressed) != 0 ||
			(GetAsyncKeyState(VKRButton)&KeyPressed) != 0 ||
			(GetAsyncKeyState(VKMButton)&KeyPressed) != 0 {
			am.lastActivityTime = time.Now()
		}
		time.Sleep(100 * time.Millisecond)
	}
}

// GetIdleSeconds 上次获取距今的空闲秒数
func (am *ActivityMonitor) GetIdleSeconds() float64 {
	return time.Since(am.lastActivityTime).Seconds()
}

// Stop 终止侦听器
func (am *ActivityMonitor) Stop() {
	am.running = false
}

// NetworkMonitor 使用WMI事件侦听网络状态
type NetworkMonitor struct {
	callback func(bool)
	running  bool
	status   bool
	stopChan chan struct{}
}

// NewNetworkMonitor 创建一个网络侦听器
func NewNetworkMonitor(callback func(bool)) *NetworkMonitor {
	return &NetworkMonitor{
		callback: callback,
		running:  true,
		status:   true, // 假设初始状态为已连接
		stopChan: make(chan struct{}),
	}
}

// Start 开始侦听网络状态
func (nm *NetworkMonitor) Start() {
	go nm.monitorNetworkEvents()
}

// monitorNetworkEvents 订阅网络变化状态
func (nm *NetworkMonitor) monitorNetworkEvents() {
	// 初始化COM
	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		fmt.Println("初始化COM失败:", err)
		return
	}
	defer ole.CoUninitialize()

	// 创建WMI连接
	unknown, err := oleutil.CreateObject("WbemScripting.SWbemLocator")
	if err != nil {
		fmt.Println("创建WMI定位器失败:", err)
		return
	}
	defer unknown.Release()

	wmi, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		fmt.Println("查询WMI接口失败:", err)
		return
	}
	defer wmi.Release()

	// 连接到WMI服务
	serviceRaw, err := oleutil.CallMethod(wmi, "ConnectServer", ".", "root\\cimv2")
	if err != nil {
		fmt.Println("连接WMI服务失败:", err)
		return
	}
	defer func(serviceRaw *ole.VARIANT) {
		_ = serviceRaw.Clear()
	}(serviceRaw)

	// 不再尝试设置安全级别，因为这可能在某些Windows版本中不支持
	// 直接继续执行后续代码

	// 由于WMI事件处理在Go中存在一些复杂性，我们改为使用轮询方式检测网络状态变化
	// 但使用更短的轮询间隔，以提高响应速度
	fmt.Println("开始监控网络状态变化...")

	// 获取初始网络状态
	currentStatus := isNetworkConnected()
	nm.status = currentStatus
	nm.callback(currentStatus)

	ticker := time.NewTicker(500 * time.Millisecond) // 每500毫秒检查一次
	defer ticker.Stop()

	for nm.running {
		select {
		case <-ticker.C:
			newStatus := isNetworkConnected()
			if newStatus != nm.status {
				nm.status = newStatus
				nm.callback(newStatus)
			}
		case <-nm.stopChan:
			return
		}
	}
}

// Stop 停止侦听器
func (nm *NetworkMonitor) Stop() {
	if nm.running {
		nm.running = false
		close(nm.stopChan)
	}
}

// isNetworkConnected 检查网路是否连接
func isNetworkConnected() bool {
	// 使用Windows API IsNetworkAlive函数
	var flags uint32
	ret, _, _ := syscall.SyscallN(
		getProcAddress("sensapi.dll", "IsNetworkAlive"),
		uintptr(unsafe.Pointer(&flags)),
		0,
		0,
	)
	return ret != 0
}

// createScheduledTask 创建Windows计划任务
func createScheduledTask(idleTime int) error {
	// 获取当前可执行文件的绝对路径
	exePath, err := os.Executable()
	if err != nil {
		return fmt.Errorf("获取可执行文件路径失败: %w", err)
	}
	exePath = filepath.Clean(exePath)

	// 创建任务名称
	taskName := "断网自动关机-柳檬科技"

	// 删除可能存在的旧任务
	deleteCmd := exec.Command("schtasks", "/Delete", "/TN", taskName, "/F")
	_ = deleteCmd.Run() // 忽略错误，因为任务可能不存在

	// 创建新任务
	// 设置为系统启动时运行，并传递空闲时间参数
	taskCommand := fmt.Sprintf("\"%s\" -i %d", exePath, idleTime)

	createCmd := exec.Command("schtasks", "/Create", "/TN", taskName, "/SC", "ONSTART",
		"/TR", taskCommand, "/RU", "SYSTEM", "/RL", "HIGHEST", "/F")

	output, err := createCmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("创建计划任务失败: %w, 输出: %s", err, string(output))
	}

	fmt.Printf("成功创建计划任务，系统启动时将自动运行（空闲时间设置为%d分钟）\n", idleTime)
	return nil
}

func main() {
	// 解析命令行参数
	var idleTime int
	var createTask bool

	pflag.IntVarP(&idleTime, "idle", "i", IdleTimeMinutes, "关机前侦听输入时长（分钟）")
	pflag.BoolVarP(&createTask, "task", "t", false, "创建Windows计划任务，系统启动时自动运行")
	pflag.Parse()

	// 如果指定了创建任务参数，则创建任务后退出
	if createTask {
		fmt.Printf("正在创建Windows计划任务（空闲时间: %d分钟）...\n", idleTime)
		err := createScheduledTask(idleTime)
		if err != nil {
			fmt.Println("创建任务失败:", err)
			fmt.Println("\n注意: 创建系统任务需要管理员权限，请尝试以管理员身份运行此程序")
			os.Exit(1)
		}
		os.Exit(0)
	}

	// Create activity monitor
	activityMonitor := NewActivityMonitor()
	activityMonitor.Start()
	defer activityMonitor.Stop()

	var shutdownTime *time.Time

	// Network status change callback
	networkStatusChanged := func(isConnected bool) {
		if !isConnected {
			fmt.Printf("[%s] 网络连接已断开\n", time.Now().Format("2006-01-02 15:04:05"))
			// Start shutdown timer when network disconnects
			t := time.Now().Add(time.Duration(idleTime) * time.Minute)
			shutdownTime = &t
		} else {
			fmt.Printf("[%s] 网络连接已恢复\n", time.Now().Format("2006-01-02 15:04:05"))
			// Cancel shutdown when network reconnects
			shutdownTime = nil
		}
	}

	// Create network monitor
	networkMonitor := NewNetworkMonitor(networkStatusChanged)
	networkMonitor.Start()
	defer networkMonitor.Stop()

	fmt.Println("===== 网络连接监控已启动 =====")
	fmt.Printf("条件满足时将自动关机:\n")
	fmt.Printf("- 网络断开后 %d 分钟内无用户输入\n", idleTime)
	fmt.Println("=============================")

	// 初始化网络状态
	initialNetworkStatus := isNetworkConnected()
	fmt.Printf("[%s] 初始网络状态: %v\n",
		time.Now().Format("2006-01-02 15:04:05"),
		func() string {
			if initialNetworkStatus {
				return "已连接"
			}
			return "已断开"
		}())

	// 如果初始状态为断开，则启动关机计时器
	if !initialNetworkStatus {
		t := time.Now().Add(time.Duration(idleTime) * time.Minute)
		shutdownTime = &t
	}

	// Handle Ctrl+C
	sigChan := make(chan os.Signal, 1)
	signal.Notify(sigChan, syscall.SIGINT, syscall.SIGTERM)

	// Main loop
	ticker := time.NewTicker(1 * time.Second)
	defer ticker.Stop()

	for {
		select {
		case <-ticker.C:
			if shutdownTime != nil {
				idleSec := activityMonitor.GetIdleSeconds()
				remainSec := time.Until(*shutdownTime).Seconds()

				// Reset timer if user activity detected
				if idleSec < 1.0 { // 检测到用户活动
					fmt.Printf("[%s] 用户活动检测，重置关机计时器\n", time.Now().Format("2006-01-02 15:04:05"))
					t := time.Now().Add(time.Duration(idleTime) * time.Minute)
					shutdownTime = &t
				} else if time.Now().After(*shutdownTime) {
					// Execute shutdown if time has elapsed
					fmt.Printf("[%s] 执行关机命令...\n", time.Now().Format("2006-01-02 15:04:05"))
					cmd := exec.Command("cmd", "/C", ShutdownCmd)
					_ = cmd.Run()
					return
				} else {
					fmt.Printf("\r[%s] 等待关机: %d秒 | 空闲: %d秒",
						time.Now().Format("2006-01-02 15:04:05"),
						int(remainSec),
						int(idleSec))
				}
			}
		case <-sigChan:
			fmt.Println("\n程序已手动终止")
			return
		}
	}
}
