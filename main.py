#!/usr/bin/env python3
import sys
import os
import argparse
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.prompt import Prompt
from rich.text import Text
import excel_processor

console = Console()


def find_excel_files():
    """Finds all valid .xlsx files in the current working directory, excluding temporary lock files."""
    files = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~") and not f.startswith(".~")]
    return sorted(files)


def display_header():
    """Displays application banner in console."""
    console.clear()
    banner = Text("ระบบผู้ช่วยบัญชีและฟาร์มสุกร (Account Assistant TUI)", style="bold white on blue", justify="center")
    console.print(Panel(banner, border_style="bright_blue", title="[bold yellow]Pig Farm Accounting System[/bold yellow]"))


def handle_menu_1():
    """Handler for Menu 1: Generate Summary Sheet."""
    console.clear()
    display_header()
    console.print(Panel("[bold cyan]เมนูที่ 1: สร้างชีทสรุปยอดสุกร (สรุปยอด)[/bold cyan]", border_style="cyan"))

    excel_files = find_excel_files()
    
    selected_file = None
    if excel_files:
        table = Table(title="พบไฟล์ Excel ในโฟลเดอร์ปัจจุบัน", show_header=True, header_style="bold green")
        table.add_column("ลำดับ", style="cyan", justify="center")
        table.add_column("ชื่อไฟล์", style="white")

        for idx, f_name in enumerate(excel_files, 1):
            table.add_row(str(idx), f_name)
        
        console.print(table)
        console.print("[dim]ป้อนลำดับไฟล์ หรือ พิมพ์ Path ของไฟล์อื่นที่ต้องการ[/dim]\n")
        
        choice = Prompt.ask("เลือกไฟล์ [1-{}] (กด Enter เพื่อเลือกไฟล์แรก)".format(len(excel_files)), default="1")
        
        if choice.isdigit():
            val = int(choice)
            if 1 <= val <= len(excel_files):
                selected_file = excel_files[val - 1]
            else:
                selected_file = choice
        else:
            selected_file = choice
    else:
        selected_file = Prompt.ask("ไม่พบไฟล์ .xlsx ในโฟลเดอร์นี้ กรุณาระบุ Path ของไฟล์ Excel")

    if not selected_file:
        console.print("[bold red]ไม่ได้เลือกไฟล์ ยกเลิกการทำงาน[/bold red]")
        Prompt.ask("\nกด Enter เพื่อกลับสู่เมนูหลัก")
        return

    selected_file = selected_file.strip()
    if not os.path.exists(selected_file):
        console.print(f"[bold red]ไม่พบไฟล์: {selected_file}[/bold red]")
        Prompt.ask("\nกด Enter เพื่อกลับสู่เมนูหลัก")
        return

    console.print(f"\n[bold yellow]กำลังประมวลผลไฟล์:[/bold yellow] [white]{selected_file}[/white]...")
    
    result = excel_processor.add_summary_sheet(selected_file)

    if result["success"]:
        success_msg = (
            f"[bold green]สร้างชีทสรุปยอดสำเร็จ![/bold green]\n\n"
            f"[white]ชื่อชีทที่ถูกสร้าง:[/white] [bold yellow]{result['sheet_name']}[/bold yellow]\n"
            f"[white]ไฟล์ที่บันทึก:[/white] [cyan]{result['file_path']}[/cyan]"
        )
        console.print(Panel(success_msg, border_style="green", title="[bold green]สำเร็จ (Success)[/bold green]"))
    else:
        console.print(Panel(f"[bold red]{result['error']}[/bold red]", border_style="red", title="[bold red]เกิดข้อผิดพลาด[/bold red]"))

    Prompt.ask("\nกด Enter เพื่อกลับสู่เมนูหลัก")


def handle_menu_2():
    """Handler for Menu 2: Generate Barn Summary Sheet (สรุปยอดแยกเล้า)."""
    console.clear()
    display_header()
    console.print(Panel("[bold cyan]เมนูที่ 2: สร้างชีทสรุปยอดสุกรแยกเล้า (สรุปยอดแยกเล้า)[/bold cyan]", border_style="cyan"))

    excel_files = find_excel_files()
    
    selected_file = None
    if excel_files:
        table = Table(title="พบไฟล์ Excel ในโฟลเดอร์ปัจจุบัน", show_header=True, header_style="bold green")
        table.add_column("ลำดับ", style="cyan", justify="center")
        table.add_column("ชื่อไฟล์", style="white")

        for idx, f_name in enumerate(excel_files, 1):
            table.add_row(str(idx), f_name)
        
        console.print(table)
        console.print("[dim]ป้อนลำดับไฟล์ หรือ พิมพ์ Path ของไฟล์อื่นที่ต้องการ[/dim]\n")
        
        choice = Prompt.ask("เลือกไฟล์ [1-{}] (กด Enter เพื่อเลือกไฟล์แรก)".format(len(excel_files)), default="1")
        
        if choice.isdigit():
            val = int(choice)
            if 1 <= val <= len(excel_files):
                selected_file = excel_files[val - 1]
            else:
                selected_file = choice
        else:
            selected_file = choice
    else:
        selected_file = Prompt.ask("ไม่พบไฟล์ .xlsx ในโฟลเดอร์นี้ กรุณาระบุ Path ของไฟล์ Excel")

    if not selected_file:
        console.print("[bold red]ไม่ได้เลือกไฟล์ ยกเลิกการทำงาน[/bold red]")
        Prompt.ask("\nกด Enter เพื่อกลับสู่เมนูหลัก")
        return

    selected_file = selected_file.strip()
    if not os.path.exists(selected_file):
        console.print(f"[bold red]ไม่พบไฟล์: {selected_file}[/bold red]")
        Prompt.ask("\nกด Enter เพื่อกลับสู่เมนูหลัก")
        return

    console.print(f"\n[bold yellow]กำลังประมวลผลสรุปยอดแยกเล้าสำหรับไฟล์:[/bold yellow] [white]{selected_file}[/white]...")
    
    result = excel_processor.add_barn_summary_sheet(selected_file)

    if result["success"]:
        success_msg = (
            f"[bold green]สร้างชีทสรุปยอดแยกเล้าสำเร็จ![/bold green]\n\n"
            f"[white]ชื่อชีทที่ถูกสร้าง:[/white] [bold yellow]{result['sheet_name']}[/bold yellow]\n"
            f"[white]ไฟล์ที่บันทึก:[/white] [cyan]{result['file_path']}[/cyan]"
        )
        console.print(Panel(success_msg, border_style="green", title="[bold green]สำเร็จ (Success)[/bold green]"))
    else:
        console.print(Panel(f"[bold red]{result['error']}[/bold red]", border_style="red", title="[bold red]เกิดข้อผิดพลาด[/bold red]"))

    Prompt.ask("\nกด Enter เพื่อกลับสู่เมนูหลัก")


def main_menu():
    """Main loop for TUI navigation."""
    while True:
        display_header()
        
        menu_table = Table(show_header=True, header_style="bold magenta", expand=True)
        menu_table.add_column("ตัวเลือก", style="bold yellow", justify="center", width=10)
        menu_table.add_column("รายการฟังก์ชันการทำงาน", style="bold white")
        
        menu_table.add_row("1", "สร้างชีทสรุปยอดสุกร (สรุปยอด)")
        menu_table.add_row("2", "สร้างชีทสรุปยอดสุกรแยกเล้า (สรุปยอดแยกเล้า)")
        menu_table.add_row("0", "ออกจากโปรแกรม (Exit)")
        
        console.print(menu_table)
        console.print()
        
        choice = Prompt.ask("เลือกเมนูการทำงาน [0-2]", choices=["0", "1", "2"], default="1")
        
        if choice == "1":
            handle_menu_1()
        elif choice == "2":
            handle_menu_2()
        elif choice == "0":
            console.print("\n[bold green]ขอบคุณที่ใช้งานโปรแกรม ขอบคุณครับ![/bold green]")
            sys.exit(0)


def main():
    parser = argparse.ArgumentParser(description="Pig Farm Accounting System (Account Assistant)")
    parser.add_argument("--tui", action="store_true", help="Launch in Terminal UI mode")
    parser.add_argument("--gui", action="store_true", help="Launch in Graphical User Interface (GUI) mode")
    
    args, unknown = parser.parse_known_args()
    
    if args.tui:
        main_menu()
    else:
        # Default mode is GUI
        try:
            import gui
            gui.launch_gui()
        except Exception as e:
            console.print(f"[bold yellow]คำเตือน: ไม่สามารถเปิดโหมด GUI ได้ ({str(e)}) ระบบจะสลับไปใช้ TUI แทน[/bold yellow]")
            main_menu()


if __name__ == "__main__":
    main()
