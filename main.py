# -*- coding: utf_8 -*-

import subprocess
import os
import sys
import ctypes
import time

def is_admin():
    """检查是否以管理员权限运行"""
    try:
        return os.getuid() == 0  # 对于 Unix/Linux
    except AttributeError:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0  # 对于 Windows

def run_as_admin():
    """以管理员权限重新运行程序，并关闭原窗口"""
    if not is_admin():
        print("当前程序未以管理员权限运行，正在尝试提权...")
        try:
            script = os.path.abspath(sys.argv[0])
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}"', None, 1)
            os._exit(0)  # 强制关闭当前进程
        except Exception as e:
            print(f"提权失败: {e}")
            input("按回车退出程序...")
            sys.exit(1)

def run_command(command):
    """运行系统命令并返回结果"""
    try:
        result = subprocess.run(command, capture_output=True, text=True, shell=True)
        return result.stdout.strip()
    except Exception as e:
        print(f"命令执行失败: {e}")
        return None

def clear_console():
    """清空控制台，兼容 Windows 和其他环境"""
    if os.name == 'nt':  # Windows
        os.system('cls')
    else:  # Unix/Linux/Mac
        os.system('clear')
    print("\n" * 100)

def open_port_menu():
    """开放端口功能"""
    while True:
        clear_console()
        print("=============")
        print("功能：开放端口")
        print("\n请输入端口号，按回车确认。")
        print("键入 q 返回首页菜单。")
        print("=============")
        port = input("请输入：").strip()
        if port.lower() == 'q':
            break
        if not port.isdigit():
            print("请输入有效的端口号！")
            input("按回车继续...")
            continue

        print(f"正在开放端口 {port} 的规则...")
        run_command(f'netsh advfirewall firewall delete rule name="开放端口 {port} 入站" protocol=TCP localport={port}')
        run_command(f'netsh advfirewall firewall delete rule name="开放端口 {port} 出站" protocol=TCP localport={port}')
        run_command(f'netsh advfirewall firewall add rule name="开放端口 {port} 入站" dir=in action=allow protocol=TCP localport={port}')
        run_command(f'netsh advfirewall firewall add rule name="开放端口 {port} 出站" dir=out action=allow protocol=TCP localport={port}')
        print(f"端口 {port} 已成功开放（入站和出站规则）！")
        input("\n按回车返回菜单...")

def close_port_menu():
    """封禁端口功能"""
    while True:
        clear_console()
        print("=============")
        print("功能：封禁端口")
        print("\n请输入端口号，按回车确认。")
        print("键入 q 返回首页菜单。")
        print("=============")
        port = input("请输入：").strip()
        if port.lower() == 'q':
            break
        if not port.isdigit():
            print("请输入有效的端口号！")
            input("按回车继续...")
            continue

        print(f"正在封禁端口 {port} 的规则...")

        # ==== 入站规则 ====
        # 先检查入站规则是否存在
        show_in_cmd = f'netsh advfirewall firewall show rule name="封禁端口 {port} 入站"'
        show_in_result = run_command(show_in_cmd)

        if not show_in_result or "No rules match" in show_in_result:
            # 说明不存在该规则，可以创建
            result_in = run_command(
                f'netsh advfirewall firewall add rule '
                f'name="封禁端口 {port} 入站" '
                f'dir=in action=block protocol=TCP localport={port}'
            )
            if "Ok" in (result_in or ""):
                print(f"【入站】规则已成功创建 -> 封禁端口 {port}")
            else:
                print(f"【入站】规则创建失败，请检查权限或配置！\n返回信息: {result_in}")
        else:
            # 已存在对应规则，可以选择提示已存在，或直接删除后再创建
            print(f"【入站】端口 {port} 的封禁规则已存在。无需重复创建。")
            # 如果想强制重新创建，可以先删除再创建：
            # run_command(f'netsh advfirewall firewall delete rule name="封禁端口 {port} 入站"')
            # 再重复上面的 add rule

        # ==== 出站规则 ====
        # 同理先检查出站规则是否存在
        show_out_cmd = f'netsh advfirewall firewall show rule name="封禁端口 {port} 出站"'
        show_out_result = run_command(show_out_cmd)

        if not show_out_result or "No rules match" in show_out_result:
            # 说明不存在该规则，可以创建
            result_out = run_command(
                f'netsh advfirewall firewall add rule '
                f'name="封禁端口 {port} 出站" '
                f'dir=out action=block protocol=TCP localport={port}'
            )
            if "Ok" in (result_out or ""):
                print(f"【出站】规则已成功创建 -> 封禁端口 {port}")
            else:
                print(f"【出站】规则创建失败，请检查权限或配置！\n返回信息: {result_out}")
        else:
            print(f"【出站】端口 {port} 的封禁规则已存在。无需重复创建。")
            # 同理，如果想强制覆盖，则先delete再add

        input("\n按回车返回菜单...")




def scan_and_delete_rules():
    """扫描指定端口的防火墙规则并选择删除"""
    while True:
        clear_console()
        print("=============")
        print("功能：删除端口规则")
        print("\n请输入要扫描的端口号，按回车确认。")
        print("键入 q 返回首页菜单。")
        print("=============")
        port = input("请输入：").strip()
        if port.lower() == 'q':
            break
        if not port.isdigit():
            print("请输入有效的端口号！")
            input("按回车继续...")
            continue

        print(f"\n正在扫描端口 {port} 的防火墙规则...\n")
        rules = []
        command = f'netsh advfirewall firewall show rule name=all'
        output = run_command(command)

        if output:
            current_rule = {}
            for line in output.splitlines():
                # 识别规则名称
                if "规则名称" in line or "Rule Name" in line:
                    if current_rule:
                        # 保存当前规则
                        if current_rule.get("port") == port:
                            rules.append(current_rule)
                        current_rule = {}
                    current_rule["name"] = line.split(":", 1)[-1].strip()
                
                # 识别端口
                if "本地端口" in line or "LocalPort" in line:
                    current_rule["port"] = line.split(":", 1)[-1].strip()

            # 添加最后一条规则
            if current_rule.get("port") == port:
                rules.append(current_rule)

        if not rules:
            print(f"未找到与端口 {port} 相关的规则。")
            input("按回车返回...")
            continue

        print("\n找到以下规则：")
        for i, rule in enumerate(rules, start=1):
            print(f"{i}. 规则名称: {rule['name']}, 本地端口: {rule['port']}")

        while True:
            selections = input("\n请输入要删除的规则序号 (用逗号分隔，输入 'q' 返回主菜单，输入 'x' 取消操作)：").strip()
            if selections.lower() == 'q':
                return  # 返回主菜单
            if selections.lower() == 'x':
                break  # 取消操作，重新等待输入端口
            try:
                indices = [int(x) for x in selections.replace('，', ',').split(",") if x.strip().isdigit()]
                for i in indices:
                    if 1 <= i <= len(rules):
                        rule_name = rules[i - 1]["name"]
                        result = run_command(f'netsh advfirewall firewall delete rule name="{rule_name}"')
                        if "已成功删除" in result or "deleted successfully" in result:
                            print(f"规则 '{rule_name}' 删除成功！")
                        else:
                            print(f"规则 '{rule_name}' 删除失败：{result}")
                    else:
                        print(f"序号 {i} 无效，跳过。")
                input("\n操作完成，按回车返回...")
                break  # 完成删除后返回端口输入阶段
            except ValueError:
                print("请输入有效的规则序号！")

def open_firewall_gui():
    """打开高级安全 Windows Defender 防火墙"""
    clear_console()
    print("正在打开高级安全 Windows Defender 防火墙...")

    # 以异步的方式打开防火墙
    subprocess.Popen("wf.msc", shell=True)
    
    # 如果想让用户看到“正在打开……”的提示信息，可以 sleep 几秒
    time.sleep(1)

    # 不再停留等待用户敲回车，直接返回主菜单
    return


def main_menu():
    """主菜单"""
    while True:
        clear_console()
        logo = r"""
    ___           __  _______      __            
   / _ \___  ____/ /_/ ___/ /_____/ /__ ____    _/|
  / ___/ _ \/ __/ __/ /__/ __/ __/ / -_) __/   > _<
 /_/   \___/_/  \__/\___/\__/_/ /_/\__/_/      |/    v₁.₀₀
                                                
  

"""
        print(logo)
        print("\n\n=======================================")
        print("功能菜单：")
        print("")
        print("1. 开放端口")
        print("2. 封禁端口")
        print("3. 删除端口规则")
        print("4. 打开[高级安全Windows Defender防火墙]")
        print("\nq. 退出程序")
        print("")
        print("by. lumosWei   ")
        print("=======================================")
        choice = input("请选择功能编号：").strip()

        
        if choice == "1":
            open_port_menu()
        elif choice == "2":
            close_port_menu()
        elif choice == "3":
            scan_and_delete_rules()
        elif choice == "4":
            open_firewall_gui()
        elif choice.lower() == 'q':
            print("程序已退出。")
            break
        else:
            print("请输入有效的选项！")
            input("按回车继续...")

if __name__ == "__main__":
    run_as_admin()  # 确保以管理员权限运行
    main_menu()
