import json
import logging
import os
import ssl
import time
import re
import mimetypes  # [新增] 用于正确识别附件类型
from dataclasses import dataclass
from email import message_from_bytes
from email.header import decode_header
from email.message import EmailMessage
from email.policy import default
from pathlib import Path
from typing import Iterable, Optional

import imaplib
import smtplib

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None


STATE_PATH = Path("state.json")


@dataclass(frozen=True)
class Config:
    src_email: str
    src_password: str
    imap_host: str
    imap_port: int
    imap_ssl: bool
    imap_folder: str
    imap_timeout: int

    smtp_user: str
    smtp_password: str
    smtp_host: str
    smtp_port: int
    smtp_ssl: bool

    dest_email: str
    poll_interval_seconds: int


def _env_bool(name: str, default: bool) -> bool:
    v = os.getenv(name)
    if v is None:
        return default
    return v.strip().lower() in {"1", "true", "yes", "y", "on"}


def _env_int(name: str, default: int) -> int:
    v = os.getenv(name)
    if v is None or not v.strip():
        return default
    return int(v)


def decode_str(s):
    """解码邮件中的编码字符串"""
    if not s: return ""
    parts = decode_header(s)
    decoded = []
    for value, charset in parts:
        if isinstance(value, bytes):
            # 尝试常见编码
            if charset:
                try:
                    value = value.decode(charset, errors='ignore')
                except LookupError:
                    # 如果字符集无法识别，尝试 utf-8 或 gb18030
                    value = value.decode('gb18030', errors='ignore')
            else:
                value = value.decode('utf-8', errors='ignore')
        decoded.append(str(value))
    return re.sub(r'\s+', ' ', "".join(decoded)).strip()


def load_config() -> Config:
    if load_dotenv:
        load_dotenv(override=False)

    cfg = Config(
        src_email=os.environ["SRC_EMAIL"],
        src_password=os.environ["SRC_PASSWORD"],
        imap_host=os.environ["IMAP_HOST"],
        imap_port=_env_int("IMAP_PORT", 993),
        imap_ssl=_env_bool("IMAP_SSL", True),
        imap_folder=os.getenv("IMAP_FOLDER", "INBOX"),
        imap_timeout=_env_int("IMAP_TIMEOUT", 120),
        smtp_user=os.environ["SMTP_USER"],
        smtp_password=os.environ["SMTP_PASSWORD"],
        smtp_host=os.environ["SMTP_HOST"],
        smtp_port=_env_int("SMTP_PORT", 465),
        smtp_ssl=_env_bool("SMTP_SSL", True),
        dest_email=os.environ["DEST_EMAIL"],
        poll_interval_seconds=_env_int("POLL_INTERVAL_SECONDS", 3600),
    )
    return cfg


def load_state() -> dict:
    if not STATE_PATH.exists():
        return {}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(state: dict) -> None:
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def imap_connect(cfg: Config):
    use_ssl = cfg.imap_ssl or cfg.imap_port == 993
    
    if use_ssl:
        print(f"Using IMAP SSL connection (Host: {cfg.imap_host}:{cfg.imap_port})")
        return imaplib.IMAP4_SSL(cfg.imap_host, cfg.imap_port, timeout=cfg.imap_timeout)
    else:
        print(f"Using IMAP plain connection (Host: {cfg.imap_host}:{cfg.imap_port})")
        return imaplib.IMAP4(cfg.imap_host, cfg.imap_port, timeout=cfg.imap_timeout)


def smtp_connect(cfg: Config) -> smtplib.SMTP:
    # [优化] 增加超时时间到 300秒，防止大附件（如PDF）发送时断开
    timeout_sec = 300 
    
    if cfg.smtp_ssl:
        context = ssl.create_default_context()
        return smtplib.SMTP_SSL(cfg.smtp_host, cfg.smtp_port, context=context, timeout=timeout_sec)
    
    server = smtplib.SMTP(cfg.smtp_host, cfg.smtp_port, timeout=timeout_sec)
    server.starttls(context=ssl.create_default_context())
    return server


def _imap_ok(resp) -> bool:
    if not resp or len(resp) < 1:
        return False
    return resp[0] == "OK"


def _ensure_selected(imap: imaplib.IMAP4, folder: str) -> None:
    sel = imap.select(folder, readonly=False)
    if not _imap_ok(sel):
        raise RuntimeError(f"IMAP select failed: {sel}")


def _build_forward_message(
    cfg: Config,
    raw_bytes: bytes,
    original_uid: str,
) -> EmailMessage:
    """构造转发邮件，包含附件（增强版：针对 PDF 和文件名修复）"""
    # 1. 解析原始邮件
    original_msg = message_from_bytes(raw_bytes, policy=default)
    
    subject = decode_str(original_msg.get('Subject'))
    from_info = decode_str(original_msg.get('From'))
    
    # 2. 创建新邮件对象
    forward_msg = EmailMessage()
    forward_msg['Subject'] = f"[转发] {subject}"
    forward_msg['From'] = cfg.smtp_user
    forward_msg['To'] = cfg.dest_email
    
    # 3. 提取正文
    body_part = original_msg.get_body(preferencelist=('html', 'plain'))
    
    notice_text = f"<p style='color:gray;font-size:12px;'>--- 原始发件人: {from_info} ---</p><hr>"
    
    if body_part:
        try:
            content = body_part.get_content()
            ctype = body_part.get_content_type()
            if ctype == 'text/html':
                forward_msg.add_alternative(notice_text + content, subtype='html')
            else:
                forward_msg.set_content(f"--- 原始发件人: {from_info} ---\n\n" + content)
        except Exception:
            forward_msg.set_content(f"--- 原始发件人: {from_info} ---\n(正文解析异常，请查看附件)")
    else:
        forward_msg.set_content(f"--- 原始发件人: {from_info} ---\n(无正文内容)")

    # 4. 遍历并处理附件 (核心修改部分)
    attachment_count = 0
    
    for part in original_msg.walk():
        # 跳过 multipart 容器
        if part.get_content_maintype() == 'multipart':
            continue
        # 跳过正文本身
        if part == body_part:
            continue
            
        # 获取文件名
        filename = part.get_filename()
        
        # 针对某些 PDF 只有 Content-Type 但没有 Content-Disposition filename 的情况
        # 或者 filename 解码失败的情况，我们手动赋予一个名字
        original_ctype = part.get_content_type().lower()
        
        # 如果没有文件名，但是类型是 PDF，强制生成一个文件名
        if not filename and 'pdf' in original_ctype:
            filename = f"document_{attachment_count}.pdf"
        
        if filename:
            # 解码文件名
            filename = decode_str(filename)
            
            # 获取附件二进制内容
            payload = part.get_payload(decode=True)
            
            if payload:
                attachment_count += 1
                
                # --- 核心修复：MIME 类型判定逻辑 ---
                # 1. 先用 mimetypes 根据后缀猜 (最准确)
                ctype, encoding = mimetypes.guess_type(filename)
                
                # 2. 如果猜不到，或者猜出来不是 PDF 但原邮件说是 PDF，则信原邮件
                if (ctype is None) or ('pdf' in original_ctype and 'pdf' not in str(ctype)):
                    # 如果原邮件明确说是 pdf，那就用 application/pdf
                    if 'pdf' in original_ctype:
                        ctype = 'application/pdf'
                    else:
                        # 否则沿用原邮件类型
                        ctype = original_ctype
                
                # 3. 兜底：如果还是空的，给默认值
                if not ctype or '/' not in ctype:
                    ctype = 'application/octet-stream'
                
                maintype, subtype = ctype.split('/', 1)
                
                try:
                    forward_msg.add_attachment(
                        payload,
                        maintype=maintype,
                        subtype=subtype,
                        filename=filename
                    )
                    logging.info(f"Attached: {filename} ({maintype}/{subtype})")
                except Exception as e:
                    logging.error(f"Failed to attach {filename}: {e}")
                    # 如果添加附件失败，不要让整个邮件发送失败，继续处理下一个附件
                    continue

    return forward_msg


def _uids_to_process(imap: imaplib.IMAP4, folder: str, last_uid: Optional[int]) -> list[int]:
    _ensure_selected(imap, folder)
    typ, data = imap.uid("search", None, "UNSEEN")

    if typ != "OK":
        raise RuntimeError(f"IMAP search failed: {(typ, data)}")

    if not data or not data[0]:
        return []
    raw = data[0].decode("utf-8", errors="ignore").strip()
    if not raw:
        return []
    return [int(x) for x in raw.split() if x.isdigit()]


def _build_forward_message_no_attachment(
    cfg: Config,
    raw_header: bytes,
    raw_body: bytes,
    original_uid: str,
) -> EmailMessage:
    """构造转发邮件，只包含头部和正文文本"""
    
    # 1. 尝试将 header 和 body 组合成一个完整的邮件对象进行解析
    # 这样 Python 就能自动识别 MIME 结构，把 base64 自动转回中文
    try:
        # 补全中间的空行以符合 RFC 822 标准
        full_msg_bytes = raw_header + b"\r\n" + raw_body
        original_msg = message_from_bytes(full_msg_bytes, policy=default)
    except Exception:
        # 如果解析失败，回退到原来的简单拼接逻辑
        original_msg = None

    # 提取头部信息
    if original_msg:
        subject = decode_str(original_msg.get('Subject'))
        from_info = decode_str(original_msg.get('From'))
    else:
        # 如果解析挂了，手动解 Header
        header_msg = message_from_bytes(raw_header, policy=default)
        subject = decode_str(header_msg.get('Subject'))
        from_info = decode_str(header_msg.get('From'))

    # 2. 提取正文 (核心修复点)
    body_content = ""
    is_html = False
    
    if original_msg:
        # 优先寻找 HTML 或 纯文本
        body_part = original_msg.get_body(preferencelist=('html', 'plain'))
        if body_part:
            try:
                # get_content() 会自动处理 base64 和 charset 解码
                body_content = body_part.get_content()
                if body_part.get_content_type() == 'text/html':
                    is_html = True
            except Exception as e:
                logging.warning(f"Failed to extract content from body part: {e}")
        
        # 如果 get_body 没找到（有时候结构很奇怪），手动遍历
        if not body_content:
            for part in original_msg.walk():
                if part.get_content_maintype() == 'text':
                    try:
                        body_content = part.get_content()
                        if part.get_content_type() == 'text/html':
                            is_html = True
                        break # 找到第一个文本就停止
                    except:
                        continue

    # 3. 如果上述解析彻底失败（兜底逻辑），才使用之前的暴力解码
    if not body_content:
        # 之前的逻辑，作为最后的救命稻草
        for enc in ['utf-8', 'gb18030', 'iso-8859-1']:
            try:
                body_content = raw_body.decode(enc)
                break
            except Exception:
                continue
        if not body_content:
            body_content = raw_body.decode('utf-8', errors='ignore')
        
        # 简单判断 html
        if "<html" in body_content.lower() or "<div" in body_content.lower():
            is_html = True

    # 4. 构造新邮件
    forward_msg = EmailMessage()
    forward_msg['Subject'] = f"[通知正文] {subject}"
    forward_msg['From'] = cfg.smtp_user
    forward_msg['To'] = cfg.dest_email
    
    # 优化提示文案
    notice_plain = f"--- 学校邮箱转发 (附件发送异常，仅提取正文) ---\n发件人: {from_info}\n主题: {subject}\n\n"
    notice_html = f"""
    <div style='background:#fff0f0;padding:12px;border:1px solid #fcc;color:#a00;margin-bottom:15px;border-radius:4px;'>
        <b>[系统提示]</b> 附件发送失败（可能文件过大或网络超时），已自动为您提取邮件正文。<br>
        <b>原始发件人:</b> {from_info}<br>
        <b>原始主题:</b> {subject}
    </div>
    <hr>
    """
    
    if is_html:
        # 如果是 HTML，把提示加在最前面
        # 注意：这里我们假设 body_content 已经是解码后的字符串
        forward_msg.add_alternative(notice_html + body_content, subtype='html')
    else:
        forward_msg.set_content(notice_plain + body_content)
    
    return forward_msg


def _imap_fetch_header_and_text(imap: imaplib.IMAP4, uid: int) -> tuple[bytes, bytes]:
    """只获取邮件头部和正文文本"""
    # 获取 HEADER 和 TEXT 部分
    typ, data = imap.uid("fetch", str(uid), "(BODY.PEEK[HEADER] BODY.PEEK[TEXT])")
    if typ != "OK":
        raise RuntimeError(f"IMAP fetch failed for uid={uid}: {(typ, data)}")
    
    raw_header = b""
    raw_body = b""
    
    # IMAP 返回的数据结构比较复杂，可能是列表嵌套元组
    # 格式通常是: [(b'uid (BODY[HEADER] {len}', b'Header Content'), b' (BODY[TEXT] {len}', b'Body Content'), b')']
    # 或者是分开的条目
    for item in data:
        if isinstance(item, tuple) and len(item) >= 2:
            key = item[0]
            val = item[1]
            if b'HEADER' in key:
                raw_header = val
            elif b'TEXT' in key:
                raw_body = val
    
    return raw_header, raw_body


def _imap_fetch_full_message(imap: imaplib.IMAP4, uid: int) -> bytes:
    typ, data = imap.uid("fetch", str(uid), "(RFC822)")
    if typ != "OK":
        raise RuntimeError(f"IMAP fetch failed for uid={uid}: {(typ, data)}")
    
    for item in data:
        if isinstance(item, tuple) and len(item) >= 2:
            return item[1] 
    raise RuntimeError("IMAP FETCH returned no message bytes")


def _imap_mark_forwarded(imap: imaplib.IMAP4, uid: int) -> None:
    imap.uid("store", str(uid), "+FLAGS", r"(\Seen)")


def process_once(cfg: Config) -> int:
    state = load_state()
    state_key = f"{cfg.src_email}:{cfg.imap_host}:{cfg.imap_folder}"
    last_uid = state.get(state_key)
    last_uid_int = int(last_uid) if isinstance(last_uid, int) or (isinstance(last_uid, str) and last_uid.isdigit()) else None

    forwarded = 0
    imap = None
    smtp = None
    
    try:
        imap = imap_connect(cfg)
        imap.login(cfg.src_email, cfg.src_password)

        uids = _uids_to_process(imap, cfg.imap_folder, last_uid_int)
        if not uids:
            logging.info("No new messages to forward.")
            return 0

        smtp = smtp_connect(cfg)
        smtp.login(cfg.smtp_user, cfg.smtp_password)

        for uid in uids:
            attempts = 0
            success = False
            while attempts < 3 and not success:
                attempts += 1
                try:
                    # 尝试1: 完整获取并转发
                    raw_bytes = _imap_fetch_full_message(imap, uid)
                    fwd = _build_forward_message(cfg, raw_bytes, original_uid=str(uid))
                    smtp.send_message(fwd)
                    
                    _imap_mark_forwarded(imap, uid)
                    forwarded += 1

                    if last_uid_int is None or uid > last_uid_int:
                        last_uid_int = uid
                        state[state_key] = last_uid_int
                        save_state(state)

                    logging.info("Forwarded uid=%s (attempt %d)", uid, attempts)
                    success = True
                    
                except (imaplib.IMAP4.abort, smtplib.SMTPException, OSError) as e:
                    # 如果是网络中断或发送错误
                    logging.warning("Error on uid=%s with attachments attempt=%d: %s", uid, attempts, e)
                    
                    # 第一次失败后，立即尝试【降级模式】：只转发正文
                    if attempts == 1:
                        logging.info("Attempting to forward without attachments (Fallback) for uid=%s", uid)
                        try:
                            # 重置连接以防连接已死
                            try: imap.logout() 
                            except: pass
                            try: smtp.quit() 
                            except: pass
                            
                            imap = imap_connect(cfg)
                            imap.login(cfg.src_email, cfg.src_password)
                            _ensure_selected(imap, cfg.imap_folder)
                            
                            smtp = smtp_connect(cfg)
                            smtp.login(cfg.smtp_user, cfg.smtp_password)
                            
                            # 获取简化内容
                            raw_header, raw_body = _imap_fetch_header_and_text(imap, uid)
                            fwd = _build_forward_message_no_attachment(cfg, raw_header, raw_body, original_uid=str(uid))
                            
                            smtp.send_message(fwd)
                            _imap_mark_forwarded(imap, uid)
                            forwarded += 1

                            if last_uid_int is None or uid > last_uid_int:
                                last_uid_int = uid
                                state[state_key] = last_uid_int
                                save_state(state)

                            logging.info("Fallback Success: Forwarded uid=%s without attachments.", uid)
                            success = True
                        except Exception as e2:
                            logging.exception("Fallback failed for uid=%s: %s", uid, e2)
                            time.sleep(2)
                    else:
                        # 超过1次尝试且不是fallback路径，简单的休眠重试
                        time.sleep(2)
                except Exception as e:
                    # 其他未知错误
                    logging.exception("Unexpected error uid=%s: %s", uid, e)
                    time.sleep(2)

            if not success:
                last_uid_int = uid
                state[state_key] = last_uid_int
                save_state(state)
                logging.warning("Skipped uid=%s after failed attempts", uid)
            
            time.sleep(3)

        return forwarded
    finally:
        try:
            if smtp: smtp.quit()
        except: pass
        try:
            if imap: imap.logout()
        except: pass


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Auto-forward mail from IMAP to another email via SMTP.")
    parser.add_argument("--once", action="store_true", help="Run one poll cycle then exit.")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    cfg = load_config()

    if args.once:
        n = process_once(cfg)
        logging.info("Done. Forwarded %d message(s).", n)
        return

    logging.info("Starting loop. Poll interval=%ss", cfg.poll_interval_seconds)
    while True:
        try:
            n = process_once(cfg)
            if n > 0:
                logging.info("Cycle done. Forwarded %d message(s).", n)
            else:
                # 只有当没有处理邮件时才打印 debug 或者是 silent
                pass 
        except Exception as e:
            logging.exception("Cycle failed: %s. Sleeping...", e)
        time.sleep(cfg.poll_interval_seconds)


if __name__ == "__main__":
    main()
