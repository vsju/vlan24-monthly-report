#!/usr/bin/env python
from db_utils import create_user
import sys
import getpass
import re

def validate_password(password):
    """
    Validate password strength.
    Requirements:
    - Minimum 8 characters
    - At least one uppercase letter
    - At least one lowercase letter
    - At least one digit
    """
    if len(password) < 8:
        return False, "비밀번호는 최소 8자 이상이어야 합니다."
    
    if not re.search(r'[A-Z]', password):
        return False, "비밀번호는 최소 1개의 대문자를 포함해야 합니다."
    
    if not re.search(r'[a-z]', password):
        return False, "비밀번호는 최소 1개의 소문자를 포함해야 합니다."
    
    if not re.search(r'\d', password):
        return False, "비밀번호는 최소 1개의 숫자를 포함해야 합니다."
    
    weak_passwords = ['password', 'admin123', '12345678', 'qwerty123', 'password1']
    if password.lower() in weak_passwords:
        return False, "너무 흔한 비밀번호입니다. 더 강력한 비밀번호를 사용하세요."
    
    return True, ""

def create_admin_user():
    """Create initial admin user with secure password requirements"""
    print("=" * 50)
    print("관리자 계정 생성")
    print("=" * 50)
    print()
    print("⚠️  보안 안내:")
    print("   - 강력한 비밀번호를 사용하세요 (최소 8자, 대소문자, 숫자 포함)")
    print("   - 다른 서비스에서 사용하는 비밀번호를 재사용하지 마세요")
    print("   - 비밀번호 관리자 사용을 권장합니다")
    print()
    
    username = ""
    while not username:
        username = input("관리자 ID: ").strip()
        if not username:
            print("❌ 사용자 ID는 필수입니다.")
    
    email = ""
    while not email:
        email = input("이메일: ").strip()
        if not email:
            print("❌ 이메일은 필수입니다.")
        elif '@' not in email:
            print("❌ 유효한 이메일 주소를 입력하세요.")
            email = ""
    
    password = ""
    while True:
        password = getpass.getpass("비밀번호: ").strip()
        if not password:
            print("❌ 비밀번호는 필수입니다.")
            continue
        
        valid, message = validate_password(password)
        if not valid:
            print(f"❌ {message}")
            continue
        
        password_confirm = getpass.getpass("비밀번호 확인: ").strip()
        if password != password_confirm:
            print("❌ 비밀번호가 일치하지 않습니다. 다시 입력하세요.")
            password = ""
            continue
        
        break
    
    full_name = ""
    while not full_name:
        full_name = input("이름: ").strip()
        if not full_name:
            print("❌ 이름은 필수입니다.")
    
    try:
        user_id = create_user(
            username=username,
            email=email,
            password=password,
            full_name=full_name,
            role='admin'
        )
        print()
        print("=" * 50)
        print(f"✅ 관리자 계정이 생성되었습니다!")
        print("=" * 50)
        print(f"   사용자명: {username}")
        print(f"   이메일: {email}")
        print(f"   이름: {full_name}")
        print()
        print("이제 웹 애플리케이션에 로그인할 수 있습니다.")
        print()
        return True
    except ValueError as e:
        print(f"❌ 오류: {str(e)}")
        return False
    except Exception as e:
        print(f"❌ 예상치 못한 오류가 발생했습니다: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = create_admin_user()
    sys.exit(0 if success else 1)
