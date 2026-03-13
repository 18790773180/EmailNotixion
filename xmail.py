"""邮件检查模块 - 同步 IMAP 实现 - 126邮箱Unsafe Login最终修复版"""
import imaplib
import email as email_stdlib
import time
import re
from datetime import datetime
from typing import Optional, List, Tuple


class EmailConfig:
    """邮件配置常量"""
    CONNECTION_TIMEOUT = 30
    NEW_EMAIL_WINDOW = 120
    DEFAULT_TEXT_NUM = 50
    SELECT_RETRY_MAX = 5
    SELECT_RETRY_DELAY = 1.0  # 增加延迟，避免服务器限流


class EmailNotifier:
    """同步邮件通知器 - 针对126邮箱Unsafe Login优化"""
    
    def __init__(self, host: str, user: str, token: str, logger=None):
        self.host = host
        self.user = user
        self.token = token
        self.last_uid: Optional[bytes] = None
        self.mail: Optional[imaplib.IMAP4] = None
        self.logger = logger
        self.text_num = EmailConfig.DEFAULT_TEXT_NUM
        self.last_successful_check: Optional[float] = None
        self.inbox_name: str = "INBOX"

    def _log(self, message: str, level: str = 'info') -> None:
        if self.logger:
            try:
                # 强化编码：捕获所有UnicodeError
                if isinstance(message, Exception):
                    message = str(message)
                message_bytes = str(message).encode('utf-8', errors='ignore')
                safe_message = message_bytes.decode('utf-8', errors='ignore')
                getattr(self.logger, level)(safe_message)
            except UnicodeError:
                # 最后fallback到ASCII
                ascii_msg = str(message).encode('ascii', errors='ignore').decode('ascii', errors='ignore')
                getattr(self.logger, level)(ascii_msg)
            except Exception:
                pass

    def cleanup(self) -> None:
        if self.mail:
            try:
                self.mail.logout()
            except Exception:
                pass
            finally:
                self.mail = None

    def _send_id_command(self, mail) -> bool:
        """增强IMAP ID命令（针对126 Unsafe Login，添加完整字段）"""
        try:
            # 扩展ID：模拟官方 + 网易要求（name, version, vendor, os, date 等）
            id_params = (
                '"name" "NetEaseMailClient"',  # 使用网易风格名称
                '"version" "2.0.0"',
                '"vendor" "EmailNotixion"',
                '"os" "Windows"',
                '"support-url" "https://mail.126.com"',
                '"date" "Mon, 1 Jan 2024 00:00:00 +0000"',  # 添加日期字段
                '"threading" "REFERENCES THREAD ordered subject not sent"'  # 额外线程支持
            )
            id_string = ' '.join(id_params)
            self._log(f"[EmailNotifier] 发送增强ID: {id_string[:100]}...", 'info')  # 只日志前100字符避免过长
            typ, data = mail.id(id_string)
            if typ == 'OK':
                self._log("[EmailNotifier] ID命令发送成功（增强版）", 'info')
                return True
            else:
                self._log(f"[EmailNotifier] 增强ID响应: {typ}, 数据: {data}", 'warning')
                # Fallback: 尝试最小ID
                min_id = '"name" "EmailClient" "version" "1.0"'
                typ, data = mail.id(min_id)
                if typ == 'OK':
                    self._log("[EmailNotifier] 最小ID发送成功", 'info')
                    return True
                self._log(f"[EmailNotifier] 最小ID也失败: {typ}", 'warning')
                return False
        except Exception as e:
            self._log(f"[EmailNotifier] ID命令异常（忽略）: {e}", 'warning')
            return False

    def _select_mailbox_with_retry(self, mailbox: str, mail) -> Tuple[str, Optional[List]]:
        """带重试的select方法（增加延迟）"""
        # 测试连接活性
        try:
            typ, _ = mail.noop()
            self._log(f"[EmailNotifier] NOOP前select测试: {typ}", 'debug')
        except Exception as e:
            self._log(f"[EmailNotifier] NOOP失败: {e}", 'warning')
        
        status, data = mail.select(mailbox)
        retry_count = 0
        while status != 'OK' and retry_count < EmailConfig.SELECT_RETRY_MAX:
            retry_count += 1
            self._log(f"[EmailNotifier] select失败 (状态: {status})，重试 {retry_count}/{EmailConfig.SELECT_RETRY_MAX} 次，延迟{EmailConfig.SELECT_RETRY_DELAY}s...", 'warning')
            time.sleep(EmailConfig.SELECT_RETRY_DELAY)
            status, data = mail.select(mailbox)
            if status == 'OK':
                self._log(f"[EmailNotifier] select重试成功 (第{retry_count}次)", 'info')
                return status, data
        
        if status != 'OK':
            error_msg = ''
            if data and isinstance(data, list) and len(data) > 0 and isinstance(data[0], bytes):
                try:
                    error_msg = data[0].decode('utf-8', errors='ignore')
                except UnicodeError:
                    error_msg = data[0].decode('ascii', errors='ignore')
            self._log(f"[EmailNotifier] select重试失败: {error_msg or 'Unknown Error'}", 'error')
            return status, data
        return status, data

    def test_connection(self) -> bool:
        test_mail = None
        try:
            # **优先SSL 993端口**（更安全，绕过Unsafe Login）
            self._log(f"[EmailNotifier] 尝试SSL 993连接到 {self.host}...", 'info')
            test_mail = imaplib.IMAP4_SSL(self.host, 993, timeout=EmailConfig.CONNECTION_TIMEOUT)
            
            # 获取服务器能力
            try:
                status, capabilities = test_mail.capability()
                self._log(f"[EmailNotifier] 服务器能力: {capabilities}", 'info')
            except Exception as e:
                self._log(f"[EmailNotifier] 获取能力失败: {e}", 'warning')
            
            # 登录（使用授权码）
            self._log(f"[EmailNotifier] 尝试登录 {self.user}...", 'info')
            test_mail.login(self.user, self.token)
            self._log("[EmailNotifier] 登录成功", 'info')
            
            # **登录后立即发送增强ID**（关键步骤）
            id_success = self._send_id_command(test_mail)
            if not id_success:
                self._log("[EmailNotifier] ID发送失败，但继续尝试select", 'warning')
            
            # 列出文件夹
            status, folder_list = test_mail.list()
            self._log(f"[EmailNotifier] 文件夹列表状态: {status}", 'info')
            
            # 查找收件箱
            self.inbox_name = self._find_inbox_name(test_mail)
            self._log(f"[EmailNotifier] 使用收件箱: {self.inbox_name}", 'info')
            
            # **带重试的select**
            status, data = self._select_mailbox_with_retry(self.inbox_name, test_mail)
            self._log(f"[EmailNotifier] 最终选择收件箱状态: {status}", 'info')
            
            if status != 'OK' and data:
                error_msg = data[0].decode('utf-8', errors='ignore') if isinstance(data[0], bytes) else str(data[0])
                self._log(f"[EmailNotifier] 选择收件箱错误详情: {error_msg}", 'error')
            
            # 备用收件箱（带重试）
            if status != 'OK':
                for alt_name in ['INBOX', 'inbox', '收件箱']:
                    if alt_name != self.inbox_name:
                        alt_status, _ = self._select_mailbox_with_retry(alt_name, test_mail)
                        if alt_status == 'OK':
                            self.inbox_name = alt_name
                            self._log(f"[EmailNotifier] 备用收件箱成功: {alt_name}", 'info')
                            status = 'OK'
                            break
            
            # 如果993失败，尝试143 + STARTTLS作为最终备用
            if status != 'OK':
                self._log("[EmailNotifier] 993失败，尝试143 STARTTLS...", 'warning')
                test_mail.close()
                test_mail = imaplib.IMAP4(self.host, 143, timeout=EmailConfig.CONNECTION_TIMEOUT)
                typ, _ = test_mail.starttls()
                if typ == 'OK':
                    test_mail.login(self.user, self.token)
                    self._send_id_command(test_mail)
                    status, _ = self._select_mailbox_with_retry(self.inbox_name, test_mail)
            
            return status == 'OK'
            
        except Exception as e:
            error_str = str(e).encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
            self._log(f"[EmailNotifier] 连接测试失败: {error_str}", 'error')
            return False
        finally:
            if test_mail:
                try:
                    test_mail.logout()
                except Exception:
                    pass

    def _find_inbox_name(self, mail) -> str:
        """查找收件箱的正确名称"""
        try:
            status, folder_list = mail.list()
            if status != 'OK':
                return "INBOX"
            
            for folder in folder_list or []:
                if isinstance(folder, bytes):
                    folder = folder.decode('utf-8', errors='ignore')
                
                if '"' in folder:
                    parts = [p.strip() for p in folder.split('"')]
                    if len(parts) >= 2:
                        folder_name = parts[1]
                        if folder_name.upper() == 'INBOX':
                            return folder_name
                
            return "INBOX"
        except Exception as e:
            self._log(f"[EmailNotifier] 查找收件箱失败: {e}", 'warning')
            return "INBOX"

    def _connect(self) -> None:
        # 与test_connection类似逻辑，确保生产连接也优先SSL
        try:
            if self.mail:
                try:
                    typ, _ = self.mail.noop()
                    if typ == 'OK':
                        status, _ = self._select_mailbox_with_retry(self.inbox_name, self.mail)
                        if status == 'OK':
                            return
                except Exception:
                    pass
            
            self.cleanup()
            
            # 优先SSL 993
            self.mail = imaplib.IMAP4_SSL(self.host, 993, timeout=EmailConfig.CONNECTION_TIMEOUT)
            self.mail.login(self.user, self.token)
            self._log("[EmailNotifier] 登录成功", 'info')
            self._send_id_command(self.mail)
            
            self.inbox_name = self._find_inbox_name(self.mail)
            status, data = self._select_mailbox_with_retry(self.inbox_name, self.mail)
            
            if status != 'OK':
                error_msg = data[0].decode('utf-8', errors='ignore') if data and isinstance(data[0], bytes) else str(data)
                self._log(f"[EmailNotifier] 选择收件箱失败: {error_msg}", 'error')
                
                for alt in ['INBOX', 'inbox']:
                    if alt != self.inbox_name:
                        alt_status, _ = self._select_mailbox_with_retry(alt, self.mail)
                        if alt_status == 'OK':
                            self.inbox_name = alt
                            status = 'OK'
                            break
                
                if status != 'OK':
                    # 备用143
                    self.cleanup()
                    self.mail = imaplib.IMAP4(self.host, 143, timeout=EmailConfig.CONNECTION_TIMEOUT)
                    self.mail.starttls()
                    self.mail.login(self.user, self.token)
                    self._send_id_command(self.mail)
                    status, _ = self._select_mailbox_with_retry(self.inbox_name, self.mail)
                
                if status != 'OK':
                    self.cleanup()
                    raise ConnectionError(f"无法选择收件箱（所有尝试失败）: {error_msg}")
                    
        except Exception as e:
            error_str = str(e).encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
            self._log(f"[EmailNotifier] 连接失败: {error_str}", 'error')
            self.cleanup()
            raise

    # 邮件解析方法保持不变（略，复制之前的 _html_to_text, _get_email_content 等）
    def _html_to_text(self, html: str) -> str:
        if not html:
            return ""
        def decode_qp(match):
            try:
                hex_str = match.group(0).replace('=', '')
                if len(hex_str) % 2 == 0:
                    return bytes.fromhex(hex_str).decode('utf-8', errors='ignore')
            except:
                pass
            return match.group(0)
        text = re.sub(r'(?:=[0-9A-F]{2})+', decode_qp, html)
        text = text.replace('=3D', '=')
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<[^>]+>', '', text)
        entities = {'&nbsp;': ' ', '&lt;': '<', '&gt;': '>', '&amp;': '&', '&quot;': '"'}
        for entity, char in entities.items():
            text = text.replace(entity, char)
        return re.sub(r'\s+', ' ', text).strip()

    def _get_email_content(self, msg) -> Tuple[str, str]:
        subject = ""
        if msg['Subject']:
            try:
                decoded = email_stdlib.header.decode_header(msg['Subject'])[0][0]
                subject = decoded.decode('utf-8', errors='ignore') if isinstance(decoded, bytes) else str(decoded)
            except Exception:
                subject = str(msg['Subject'])
        if len(subject) > self.text_num:
            subject = subject[:self.text_num] + "..."
        content = "（无文本内容）"
        if msg.is_multipart():
            text_content = html_content = None
            for part in msg.walk():
                ct = part.get_content_type()
                try:
                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue
                    charset = part.get_content_charset() or 'utf-8'
                    decoded = payload.decode(charset, errors='ignore')
                    if ct == "text/plain":
                        text_content = decoded
                        break
                    elif ct == "text/html" and not html_content:
                        html_content = decoded
                except Exception:
                    continue
            if text_content:
                content = self._process_content(text_content)
            elif html_content:
                content = self._process_content(self._html_to_text(html_content))
        else:
            try:
                payload = msg.get_payload(decode=True)
                if payload:
                    charset = msg.get_content_charset() or 'utf-8'
                    text = payload.decode(charset, errors='ignore')
                    if msg.get_content_type() == "text/html":
                        text = self._html_to_text(text)
                    content = self._process_content(text)
            except Exception:
                pass
        return subject, content

    def _process_content(self, text: str) -> str:
        if not text:
            return "（无文本内容）"
        text = ' '.join(line.strip() for line in text.splitlines() if line.strip())
        text = re.sub(r'\s+', ' ', text)
        if len(text) > self.text_num:
            text = text[:self.text_num] + "..."
        return text.strip() or "（无文本内容）"

    def _is_recent(self, email_time: Optional[datetime]) -> bool:
        if not email_time:
            return True
        try:
            return (time.time() - email_time.timestamp()) < EmailConfig.NEW_EMAIL_WINDOW
        except Exception:
            return True

    def check_and_notify(self) -> Optional[List[Tuple[Optional[datetime], str, str]]]:
        try:
            self._connect()
            typ, data = self.mail.uid('SEARCH', None, 'ALL')
            if typ != 'OK' or not data or not data[0]:
                typ, data = self.mail.search(None, 'ALL')
                if typ != 'OK' or not data or not data[0]:
                    return None
            all_uids = data[0].split()
            if not all_uids:
                return None
            if self.last_uid is None:
                self.last_uid = all_uids[-1]
                self.last_successful_check = time.time()
                self._log(f"[EmailNotifier] 初始化基准 UID: {self.last_uid}")
                return None
            new_emails = []
            new_last_uid = self.last_uid
            for uid in all_uids:
                if uid <= self.last_uid:
                    continue
                info = self._get_email_info(uid)
                if info and self._is_recent(info[0]):
                    new_emails.append(info)
                if uid > new_last_uid:
                    new_last_uid = uid
            self.last_uid = new_last_uid
            self.last_successful_check = time.time()
            return new_emails if new_emails else None
        except Exception as e:
            error_str = str(e).encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
            self._log(f"[EmailNotifier] 检查错误: {error_str}", 'error')
            self.cleanup()
            return None

    def _get_email_info(self, uid: bytes) -> Optional[Tuple[Optional[datetime], str, str]]:
        try:
            typ, msg_data = self.mail.uid('FETCH', uid, '(RFC822)')
            if typ != 'OK' or not msg_data or not msg_data[0]:
                return None
            msg = email_stdlib.message_from_bytes(msg_data[0][1])
            local_date = None
            try:
                date_tuple = email_stdlib.utils.parsedate_tz(msg['Date'])
                if date_tuple:
                    local_date = datetime.fromtimestamp(email_stdlib.utils.mktime_tz(date_tuple))
            except Exception:
                pass
            subject, content = self._get_email_content(msg)
            return local_date, subject, content
        except Exception as e:
            self._log(f"[EmailNotifier] 获取邮件失败 UID {uid}: {e}", 'error')
            return None
