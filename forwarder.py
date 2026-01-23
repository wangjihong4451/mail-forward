import os
import imaplib
import smtplib
import ssl
import mimetypes
from email import message_from_bytes, policy
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from datetime import datetime


# =======================
# 基础配置读取
# =======================

def load_config():
    return {
        'imap_host': os.environ['IMAP_HOST'],
        'imap_port': int(os.environ.get('IMAP_PORT', 993)),
        'imap_ssl': os.environ.get('IMAP_SSL', 'true').lower() == 'true',
        'imap_folder': os.environ.get('IMAP_FOLDER', 'INBOX'),
        'src_email': os.environ['SRC_EMAIL'],
        'src_password': os.environ['SRC_PASSWORD'],
        'dst_email': os.environ['DST_EMAIL'],
        'smtp_host': os.environ.get('SMTP_HOST', 'smtp.qq.com'),
        'smtp_port': int(os.environ.get('SMTP_PORT', 465)),
    }


# =======================
# IMAP
# =======================

def imap_login(cfg):
    if cfg['imap_ssl']:
        imap = imaplib.IMAP4_SSL(cfg['imap_host'], cfg['imap_port'])
    else:
        imap = imaplib.IMAP4(cfg['imap_host'], cfg['imap_port'])

    imap.login(cfg['src_email'], cfg['src_password'])
    imap.select(cfg['imap_folder'])
    return imap


def fetch_full_message(imap, uid):
    typ, data = imap.fetch(uid, '(RFC822)')
    if typ != 'OK':
        raise RuntimeError('Failed to fetch RFC822')
    return data[0][1]


# =======================
# 构造转发邮件（含附件）
# =======================

def build_forward_message(cfg, raw_bytes, original_uid):
    original = message_from_bytes(raw_bytes, policy=policy.default)

    fwd = EmailMessage()
    fwd['From'] = cfg['dst_email']
    fwd['To'] = cfg['dst_email']
    fwd['Subject'] = 'FWD: ' + (original.get('Subject', ''))
    fwd['Date'] = formatdate(localtime=True)
    fwd['Message-ID'] = make_msgid()

    # 复制正文
    has_body = False
    if original.is_multipart():
        for part in original.walk():
            ctype = part.get_content_type()
            disp = part.get_content_disposition()

            if ctype == 'text/plain' and disp != 'attachment':
                fwd.set_content(
                    part.get_content(),
                    charset=part.get_content_charset() or 'utf-8'
                )
                has_body = True
                break

        for part in original.walk():
            ctype = part.get_content_type()
            disp = part.get_content_disposition()

            if ctype == 'text/html' and disp != 'attachment':
                fwd.add_alternative(
                    part.get_content(),
                    subtype='html',
                    charset=part.get_content_charset() or 'utf-8'
                )
                has_body = True
                break
    else:
        fwd.set_content(original.get_content())
        has_body = True

    if not has_body:
        fwd.set_content('(正文为空，请登录邮箱查看)')

    # 复制附件
    for part in original.walk():
        disp = part.get_content_disposition()
        if disp != 'attachment':
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        filename = part.get_filename()
        if not filename:
            ext = mimetypes.guess_extension(part.get_content_type()) or '.bin'
            filename = f'attachment_{original_uid}{ext}'

        fwd.add_attachment(
            payload,
            maintype=part.get_content_maintype(),
            subtype=part.get_content_subtype(),
            filename=filename
        )

    return fwd


# =======================
# 构造转发邮件（无附件，降级）
# =======================

def build_forward_message_without_attachments(cfg, original):
    fwd = EmailMessage()
    fwd['From'] = cfg['dst_email']
    fwd['To'] = cfg['dst_email']
    fwd['Subject'] = 'FWD: ' + (original.get('Subject', ''))
    fwd['Date'] = formatdate(localtime=True)
    fwd['Message-ID'] = make_msgid()

    has_body = False

    if original.is_multipart():
        for part in original.walk():
            ctype = part.get_content_type()
            disp = part.get_content_disposition()

            if disp == 'attachment':
                continue

            if ctype == 'text/plain':
                fwd.set_content(
                    part.get_content(),
                    charset=part.get_content_charset() or 'utf-8'
                )
                has_body = True
                break

        for part in original.walk():
            ctype = part.get_content_type()
            disp = part.get_content_disposition()

            if disp == 'attachment':
                continue

            if ctype == 'text/html':
                fwd.add_alternative(
                    part.get_content(),
                    subtype='html',
                    charset=part.get_content_charset() or 'utf-8'
                )
                has_body = True
                break
    else:
        fwd.set_content(original.get_content())
        has_body = True

    if not has_body:
        fwd.set_content('(正文为空，请登录邮箱查看)')

    return fwd


# =======================
# SMTP 发送
# =======================

def send_mail(cfg, msg):
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(cfg['smtp_host'], cfg['smtp_port'], context=context) as smtp:
        smtp.login(cfg['dst_email'], cfg['src_password'])
        smtp.send_message(msg)


# =======================
# 主流程
# =======================

def process_once(cfg):
    imap = imap_login(cfg)

    typ, data = imap.search(None, 'UNSEEN')
    if typ != 'OK':
        return 0

    uids = data[0].split()
    count = 0

    for uid in uids:
        try:
            raw = fetch_full_message(imap, uid)
            fwd = build_forward_message(cfg, raw, uid.decode())
        except Exception as e:
            raw = fetch_full_message(imap, uid)
            original = message_from_bytes(raw, policy=policy.default)
            fwd = build_forward_message_without_attachments(cfg, original)

        send_mail(cfg, fwd)
        imap.store(uid, '+FLAGS', '\\Seen')
        count += 1

    imap.logout()
    return count


def main():
    cfg = load_config()
    n = process_once(cfg)
    print(f'[{datetime.now()}] Forwarded {n} mails')


if __name__ == '__main__':
    main()
