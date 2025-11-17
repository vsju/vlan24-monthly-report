# First Time Setup

## Creating Your First Admin Account

Before you can use the PowerPoint Automation Tool, you need to create an administrator account.

### Step 1: Run the Admin Creation Script

```bash
python create_admin.py
```

### Step 2: Enter Your Credentials

The script will prompt you for:
- **Username**: Your admin username (recommended: your actual name or email)
- **Email**: Your email address
- **Password**: A **strong, unique password**
  - Minimum 12 characters
  - Mix of letters, numbers, and symbols
  - Never reuse passwords from other services
- **Full Name**: Your display name

### Step 3: Login

1. Open the web application
2. Login with the credentials you just created
3. You're ready to use the tool!

## Security Best Practices

⚠️ **IMPORTANT SECURITY GUIDELINES:**

1. **Never use default or weak passwords**
   - Avoid: "admin", "password", "123456", etc.
   - Use a password manager to generate strong passwords

2. **Change passwords regularly**
   - Especially if multiple people have access
   - After any security incident

3. **Limit admin accounts**
   - Only create admin accounts for people who need full system access
   - Use regular user accounts for day-to-day work

4. **Monitor user activity**
   - Check the "작업 이력" tab regularly
   - Review who is creating reports

5. **Keep your database credentials secure**
   - Never share DATABASE_URL or other secrets
   - These are automatically managed by Replit Secrets

## Need Help?

If you encounter any issues during setup, check:
- The database is running (it should start automatically in Replit)
- You have the required packages installed (check requirements.txt)
- File permissions are correct

For more information, see replit.md
