#!/usr/bin/env python3
"""
Startup verification script for WRLD Document Generator
Checks environment variables and Google Drive connectivity before launching the app
"""

import os
import json
import sys

def check_environment():
    """Verify all required environment variables are set"""
    print("🔍 Checking environment variables...")
    
    required_vars = {
        'GOOGLE_CREDENTIALS_JSON': 'Google service account credentials',
        'GOOGLE_DRIVE_TEMPLATES_FOLDER': 'Google Drive templates folder ID'
    }
    
    missing = []
    
    for var, description in required_vars.items():
        if var not in os.environ or not os.environ[var].strip():
            missing.append(f"  ❌ {var}: {description}")
        else:
            print(f"  ✅ {var}: Set")
    
    if missing:
        print("\n⚠️  Missing environment variables:\n")
        for msg in missing:
            print(msg)
        return False
    
    return True

def check_credentials_json():
    """Verify Google credentials JSON is valid"""
    print("\n🔍 Validating Google credentials JSON...")
    
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON', '')
    
    if not creds_json:
        print("  ❌ GOOGLE_CREDENTIALS_JSON is empty")
        return False
    
    try:
        creds = json.loads(creds_json)
        required_fields = ['type', 'project_id', 'private_key', 'client_email']
        
        for field in required_fields:
            if field not in creds:
                print(f"  ❌ Missing field in credentials: {field}")
                return False
        
        if creds['type'] != 'service_account':
            print(f"  ❌ Invalid type: {creds['type']} (expected 'service_account')")
            return False
        
        print(f"  ✅ Credentials valid")
        print(f"  ✅ Service account: {creds.get('client_email', 'unknown')}")
        print(f"  ✅ Project: {creds.get('project_id', 'unknown')}")
        return True
    
    except json.JSONDecodeError as e:
        print(f"  ❌ Invalid JSON: {str(e)}")
        print("  💡 Ensure you're pasting the entire JSON as a single string")
        return False
    except Exception as e:
        print(f"  ❌ Error: {str(e)}")
        return False

def check_dependencies():
    """Verify all required Python packages are installed"""
    print("\n🔍 Checking Python dependencies...")
    
    dependencies = {
        'flask': 'Flask web framework',
        'docx': 'python-docx for Word documents',
        'google': 'google-auth for Drive API',
    }
    
    missing = []
    
    for module, description in dependencies.items():
        try:
            __import__(module)
            print(f"  ✅ {module}: {description}")
        except ImportError:
            missing.append(f"  ❌ {module}: {description}")
    
    if missing:
        print("\n⚠️  Missing dependencies:\n")
        for msg in missing:
            print(msg)
        print("\n💡 Run: pip install -r requirements.txt")
        return False
    
    return True

def main():
    """Run all checks"""
    print("=" * 60)
    print("WRLD Document Generator - Startup Verification")
    print("=" * 60)
    
    checks = [
        ("Environment Variables", check_environment),
        ("Google Credentials", check_credentials_json),
        ("Dependencies", check_dependencies),
    ]
    
    results = []
    
    for check_name, check_func in checks:
        try:
            result = check_func()
            results.append((check_name, result))
        except Exception as e:
            print(f"\n❌ {check_name} check failed: {str(e)}")
            results.append((check_name, False))
    
    print("\n" + "=" * 60)
    print("Verification Summary:")
    print("=" * 60)
    
    all_passed = True
    for check_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{status} - {check_name}")
        if not result:
            all_passed = False
    
    print("=" * 60)
    
    if all_passed:
        print("\n✅ All checks passed! Ready to start the service.\n")
        return 0
    else:
        print("\n⚠️  Some checks failed. Please fix the issues above.\n")
        return 1

if __name__ == '__main__':
    sys.exit(main())
