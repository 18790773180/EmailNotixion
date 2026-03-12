"""邮件检查模块 - 同步 IMAP 实现 - 126/163邮箱ID命令修复版"""
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
    IMAP_PORT = 143  # 使用143端口 + STARTTLS，符合官方示例


class EmailNotifier:
    """同步邮件通知器 - 针对126/163邮箱优化，添加ID命令"""
    
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
                # 修复编码错误：确保消息UTF-8兼容
                message = str(message).encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
                getattr(self.logger, level)(message)
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
        """发送IMAP ID命令（模拟Java示例）"""
        try:
            # 模拟Java HashMap IAM的键值对
            id_params = (
                '"name" "pyIMAPClient"',
                '"version" "1.0.0"',
                '"vendor" "PythonClient"',
                '"support-email" "support@example.com"'
            )
            # 构建ID命令字符串：ID (key value key value ...)
            id_string = ' '.join(id_params)
            typ, data = mail.id(id_string)
            if typ == 'OK':
                self._log(f"[EmailNotifier] ID命令发送成功: {id_string}", 'info')
                return True
            else:
                self._log(f"[EmailNotifier] ID命令响应: {typ}, 数据: {data}", 'warning')
                return False
        except Exception as e:
            self._log(f"[EmailNotifier] ID命令失败（可忽略）: {e}", 'warning')
            return False

    def test_connection(self) -> bool:
        test_mail = None
        try:
            # 使用官方推荐：IMAP4 on 143端口
            test_mail = imaplib.IMAP4(self.host, EmailConfig.IMAP_PORT, timeout=EmailConfig.CONNECTION_TIMEOUT)
            
            # 获取服务器能力
            try:
                status, capabilities = test_mail.capability()
                self._log(f"[EmailNotifier] 服务器能力: {capabilities}", 'info')
            except Exception as e:
                self._log(f"[EmailNotifier] 获取能力失败: {e}", 'warning')
            
            # 升级到TLS（STARTTLS）
            try:
                typ, data = test_mail.starttls()
                self._log(f"[EmailNotifier] STARTTLS状态: {typ}", 'info')
                if typ != 'OK':
                    raise Exception("STARTTLS失败")
            except Exception as e:
                self._log(f"[EmailNotifier] STARTTLS失败: {e}，尝试SSL连接", 'warning')
                # 备用：直接SSL on 993
                test_mail.close()
                test_mail = imaplib.IMAP4_SSL(self.host, 993, timeout=EmailConfig.CONNECTION_TIMEOUT)
            
            # 登录（使用授权码）
            self._log(f"[EmailNotifier] 尝试登录 {self.user}...", 'info')
            test_mail.login(self.user, self.token)
            self._log(f"[EmailNotifier] 登录成功", 'info')
            
            # **关键：登录后发送ID命令**（符合Java示例）
            self._send_id_command(test_mail)
            
            # 列出所有文件夹
            status, folder_list = test_mail.list()
            self._log(f"[EmailNotifier] 文件夹列表状态: {status}", 'info')
            if folder_list:
                self._log(f"[EmailNotifier] 文件夹示例: {folder_list[:2]}...", 'info')
            
            # 查找收件箱名称
            self.inbox_name = self._find_inbox_name(test_mail)
            self._log(f"[EmailNotifier] 使用收件箱: {self.inbox_name}", 'info')
            
            # 选择收件箱（现在应该成功，避免Unsafe Login）
            status, data = test_mail.select(self.inbox_name)
            self._log(f"[EmailNotifier] 选择收件箱状态: {status}", 'info')
            if status != 'OK' and data:
                error_msg = data[0].decode('utf-8', errors='ignore') if isinstance(data[0], bytes) else str(data[0])
                self._log(f"[EmailNotifier] 选择收件箱错误详情: {error_msg}", 'error')
            
            # 尝试备用收件箱名称
            if status != 'OK':
                for alt_name in ['INBOX', 'inbox', '收件箱']:
                    if alt_name != self.inbox_name:
                        alt_status, _ = test_mail.select(alt_name)
                        if alt_status == 'OK':
                            self.inbox_name = alt_name
                            self._log(f"[EmailNotifier] 备用收件箱成功: {alt_name}", 'info')
                            status = 'OK'
                            break
            
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
                
                # 解析文件夹名称（格式如：() "/" "INBOX"）
                if '"' in folder:
                    parts = [p.strip() for p in folder.split('"')]
                    if len(parts) >= 2:
                        folder_name = parts[1]  # 取引号内的名称
                        if folder_name.upper() == 'INBOX':
                            return folder_name
                
            return "INBOX"
        except Exception as e:
            self._log(f"[EmailNotifier] 查找收件箱失败: {e}", 'warning')
            return "INBOX"

    def _connect(self) -> None:
        try:
            # 检查现有连接
            if self.mail:
                try:
                    typ, _ = self.mail.noop()
                    if typ == 'OK':
                        # 重新选择收件箱
                        status, _ = self.mail.select(self.inbox_name)
                        if status == 'OK':
                            return
                except Exception:
                    pass
            
            self.cleanup()
            
            # 新连接：IMAP4 on 143 + STARTTLS
            try:
                self.mail = imaplib.IMAP4(self.host, EmailConfig.IMAP_PORT, timeout=EmailConfig.CONNECTION_TIMEOUT)
                typ, _ = self.mail.starttls()
                if typ != 'OK':
                    raise Exception("STARTTLS失败")
            except Exception as e:
                self._log(f"[EmailNotifier] STARTTLS失败: {e}，切换SSL", 'warning')
                self.mail = imaplib.IMAP4_SSL(self.host, 993, timeout=EmailConfig.CONNECTION_TIMEOUT)
            
            # 登录
            self.mail.login(self.user, self.token)
            self._log(f"[EmailNotifier] 登录成功", 'info')
            
            # **关键：登录后发送ID命令**
            self._send_id_command(self.mail)
            
            # 查找并选择收件箱
            self.inbox_name = self._find_inbox_name(self.mail)
            status, data = self.mail.select(self.inbox_name)
            
            if status != 'OK':
                error_msg = data[0].decode('utf-8', errors='ignore') if data and isinstance(data[0], bytes) else str(data)
                self._log(f"[EmailNotifier] 选择收件箱失败: {error_msg}", 'error')
                
                # 备用尝试
                for alt in ['INBOX', 'inbox']:
                    if alt != self.inbox_name:
                        status, _ = self.mail.select(alt)
                        if status == 'OK':
                            self.inbox_name = alt
                            break
                
                if status != 'OK':
                    self.cleanup()
                    raise ConnectionError(f"无法选择收件箱: {status} - {error_msg}")
                    
        except Exception as e:
            error_str = str(e).encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
            self._log(f"[EmailNotifier] 连接失败: {error_str}", 'error')
            self.cleanup()
            raise

    # 以下方法保持不变（邮件解析相关）
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
            
            # 搜索邮件（使用UID以提高效率）
            typ, data = self.mail.uid('SEARCH', None, 'ALL')
            if typ != 'OK' or not data or not data[0]:
                typ, data = self.mail.search(None, 'ALL')  # 备用普通搜索
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