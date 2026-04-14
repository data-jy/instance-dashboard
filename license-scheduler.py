#!/usr/bin/env python3
import os
import json
import sys
from datetime import datetime, timedelta

try:
    import requests
except ImportError:
    requests = None

DATA_FILE = '/home/rocky/license-data.json'

def daysLeft(expiry_str):
    """Calculate days left from expiry date"""
    exp_date = datetime.strptime(expiry_str, '%Y-%m-%d')
    now = datetime.now()
    now = now.replace(hour=0, minute=0, second=0, microsecond=0)
    exp_date = exp_date.replace(hour=0, minute=0, second=0, microsecond=0)
    delta = exp_date - now
    return delta.days

def statusOf(days):
    if days < 0:
        return 'expired'
    if days <= 7:
        return 'danger'
    if days <= 30:
        return 'warning'
    if days <= 90:
        return 'caution'
    return 'ok'

def should_notify(license_obj, settings):
    """Check if this license should trigger notification"""
    dl = daysLeft(license_obj.get('expiryDate', ''))
    status = statusOf(dl)
    
    # Check if within alert range (90 days)
    if dl > 90:
        return False
    
    # Check if already notified for this day
    sent_alerts = license_obj.get('sentAlerts', [])
    today = datetime.now().strftime('%Y-%m-%d')
    if today in sent_alerts:
        return False
    
    return True

def send_webhook(webhook_url, payload):
    """Send webhook notification"""
    if not requests:
        print(f"[SKIP] requests module not available: {payload[:50]}...")
        return False
    
    try:
        # Use rocky server proxy for CORS
        res = requests.post('http://localhost:5002/webhook-proxy', json={
            'url': webhook_url,
            'payload': {'botName': 'License Monitor', 'botIconImage': '', 'text': payload}
        }, timeout=10)
        data = res.json()
        return data.get('ok', False)
    except Exception as e:
        print(f"[ERROR] Webhook failed: {e}")
        return False

def build_message(license_obj, days_left, emoji):
    """Build notification message"""
    if days_left < 0:
        status_text = f"만료 {abs(days_left)}일 경과"
    elif days_left == 0:
        status_text = "오늘 만료!"
    else:
        status_text = f"{days_left}일 후 만료"
    
    msg = f"{emoji} [라이센스 만료 알림]\n"
    msg += f"• 벤더/제품: {license_obj.get('vendor', '-')} - {license_obj.get('product', '-')}\n"
    msg += f"• 상태: `{status_text}`\n"
    msg += f"• 만료일: {license_obj.get('expiryDate', '-')}\n"
    msg += f"• 망구분: {license_obj.get('net', '-')}\n"
    msg += f"• 환경: {license_obj.get('env', '-')}\n"
    if license_obj.get('note'):
        msg += f"• 비고: {license_obj['note']}\n"
    
    return msg

def main():
    if not os.path.exists(DATA_FILE):
        print(f"[INFO] Data file not found: {DATA_FILE}")
        return
    
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"[ERROR] Failed to read data file: {e}")
        return
    
    licenses = data.get('licenses', [])
    settings = data.get('settings', {})
    webhook_url = settings.get('webhookUrl', '')
    
    if not webhook_url:
        print("[SKIP] No webhook URL configured")
        return
    
    notified_count = 0
    for license_obj in licenses:
        if not should_notify(license_obj, settings):
            continue
        
        dl = daysLeft(license_obj.get('expiryDate', ''))
        status = statusOf(dl)
        
        emoji_map = {
            'expired': '🚨',
            'danger': '🔴',
            'warning': '🟠',
            'caution': '🟡',
            'ok': '🟢'
        }
        emoji = emoji_map.get(status, '🟡')
        
        msg = build_message(license_obj, dl, emoji)
        
        print(f"[NOTIFY] {license_obj.get('vendor', 'Unknown')} - {license_obj.get('product', 'Unknown')} ({dl}days)")
        
        ok = send_webhook(webhook_url, msg)
        
        if ok:
            # Mark as notified today
            if 'sentAlerts' not in license_obj:
                license_obj['sentAlerts'] = []
            today = datetime.now().strftime('%Y-%m-%d')
            if today not in license_obj['sentAlerts']:
                license_obj['sentAlerts'].append(today)
            notified_count += 1
            print(f"  → Success")
        else:
            print(f"  → Failed")
    
    # Save updated data
    if notified_count > 0:
        try:
            with open(DATA_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"[DONE] Notified {notified_count} licenses, data updated")
        except Exception as e:
            print(f"[ERROR] Failed to save data: {e}")
    else:
        print("[DONE] No licenses to notify")

if __name__ == '__main__':
    main()
