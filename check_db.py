import sqlite3
import os

p = r'c:\antigravity\navercafe_auto\chrome_profiles\zerooo007\Default\Network\Cookies'
print(f"Exists: {os.path.exists(p)}")
if os.path.exists(p):
    try:
        conn = sqlite3.connect(p)
        rows = [row for row in conn.execute('SELECT host_key, name FROM cookies WHERE name="NID_SES"')]
        print(f"NID_SES cookies: {rows}")
        
        all_cookies = [row for row in conn.execute('SELECT host_key, name FROM cookies WHERE host_key LIKE "%naver.com%"')]
        print(f"All naver cookies count: {len(all_cookies)}")
    except Exception as e:
        print(f"Error: {e}")
