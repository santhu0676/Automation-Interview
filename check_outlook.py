"""
Outlook Connection Diagnostic Tool
Run this to check if Outlook is properly configured
"""
import sys

def check_outlook():
    """Check Outlook installation and configuration"""
    print("=" * 70)
    print("  OUTLOOK CONNECTION DIAGNOSTIC")
    print("=" * 70)
    print()
    
    # Check 1: pywin32 installed
    print("1️⃣ Checking pywin32 installation...")
    try:
        import win32com.client
        print("   ✓ pywin32 is installed")
    except ImportError:
        print("   ✗ pywin32 is NOT installed")
        print("   Fix: Run 'pip install pywin32'")
        return False
    
    # Check 2: Outlook installed
    print("\n2️⃣ Checking Outlook installation...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("   ✓ Outlook is installed")
    except Exception as e:
        print(f"   ✗ Outlook is NOT installed or not accessible")
        print(f"   Error: {str(e)}")
        print("   Fix: Install Microsoft Outlook (desktop version)")
        return False
    
    # Check 3: Outlook accounts configured
    print("\n3️⃣ Checking Outlook email accounts...")
    try:
        namespace = outlook.GetNamespace("MAPI")
        accounts = namespace.Accounts
        
        if accounts.Count == 0:
            print("   ✗ No email accounts configured in Outlook")
            print("   Fix: Open Outlook and add an email account")
            return False
        else:
            print(f"   ✓ {accounts.Count} email account(s) configured:")
            for i in range(1, min(accounts.Count + 1, 4)):  # Show max 3 accounts
                account = accounts.Item(i)
                print(f"      - {account.DisplayName}")
    except Exception as e:
        print(f"   ✗ Cannot access Outlook accounts")
        print(f"   Error: {str(e)}")
        return False
    
    # Check 4: Can create email
    print("\n4️⃣ Testing email creation...")
    try:
        mail = outlook.CreateItem(0)  # 0 = MailItem
        mail.Subject = "Test - Connection Successful"
        print("   ✓ Can create email items")
        print("   ✓ All checks passed!")
    except Exception as e:
        print(f"   ✗ Cannot create email items")
        print(f"   Error: {str(e)}")
        return False
    
    print("\n" + "=" * 70)
    print("  ✅ OUTLOOK IS READY TO USE!")
    print("=" * 70)
    print()
    print("You can now run the main script to send interview notifications.")
    return True


if __name__ == "__main__":
    try:
        success = check_outlook()
        if not success:
            print("\n" + "=" * 70)
            print("  ⚠️  OUTLOOK IS NOT READY")
            print("=" * 70)
            print("\nPlease fix the issues above and run this diagnostic again.")
            sys.exit(1)
    except Exception as e:
        print(f"\n✗ Unexpected error: {str(e)}")
        sys.exit(1)
