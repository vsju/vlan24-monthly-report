#!/usr/bin/env python3
import os
import sys
import config

def show_menu():
    print("\n" + "="*50)
    print("PowerPoint 자동화 도구")
    print("="*50)
    print("1. 이미지 삽입 (Step 1)")
    print("2. Grafana 통계 삽입 (Step 2)")
    print("3. 전체 프로세스 실행 (Step 1 + Step 2)")
    print("4. 디렉토리 구조 생성")
    print("5. 환경 설정")
    print("6. 종료")
    print("="*50)

def create_directories():
    dirs = [
        config.BASE_TEMPLATE_DIR,
        config.BASE_IMAGE_DIR,
        config.OUTPUT_DIR_WITH_IMAGES,
        config.OUTPUT_DIR
    ]
    
    for directory in dirs:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"Created: {directory}")
        else:
            print(f"Exists: {directory}")
    
    print("\nDirectory structure created successfully!")

def show_config():
    print("\n현재 설정:")
    print(f"Template Directory: {config.BASE_TEMPLATE_DIR}")
    print(f"Image Directory: {config.BASE_IMAGE_DIR}")
    print(f"Intermediate Output: {config.OUTPUT_DIR_WITH_IMAGES}")
    print(f"Final Output: {config.OUTPUT_DIR}")
    print(f"Grafana URL: {config.GRAFANA_URL}")
    print(f"API Key: {'*' * 20 if config.API_KEY else 'Not Set'}")
    print(f"\nCustomers configured: {len(config.DASHBOARD_MAP)}")

def run_image_insertion():
    print("\nRunning image insertion script...")
    os.system("python insert_images.py")

def run_stats_insertion(customer=None):
    print("\nRunning Grafana statistics insertion script...")
    if customer:
        os.system(f"python numinsert3.py {customer}")
    else:
        os.system("python numinsert3.py")

def run_full_process():
    print("\nRunning full process (images + statistics)...")
    run_image_insertion()
    print("\n" + "="*50)
    run_stats_insertion()

def main():
    while True:
        show_menu()
        choice = input("\nSelect option (1-6): ").strip()
        
        if choice == "1":
            run_image_insertion()
        elif choice == "2":
            customer = input("Enter customer name (press Enter for all): ").strip()
            run_stats_insertion(customer if customer else None)
        elif choice == "3":
            run_full_process()
        elif choice == "4":
            create_directories()
        elif choice == "5":
            show_config()
        elif choice == "6":
            print("Exiting...")
            sys.exit(0)
        else:
            print("Invalid option. Please try again.")
        
        input("\nPress Enter to continue...")

if __name__ == "__main__":
    print("Welcome to PowerPoint Automation Tool")
    print("This tool helps automate PowerPoint report generation with images and Grafana statistics.")
    
    if not os.path.exists(config.BASE_TEMPLATE_DIR):
        print(f"\nWarning: Template directory not found at {config.BASE_TEMPLATE_DIR}")
        create = input("Would you like to create the directory structure now? (y/n): ").strip().lower()
        if create == 'y':
            create_directories()
    
    main()
