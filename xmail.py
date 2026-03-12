"""邮件检查模块 - 同步 IMAP 实现 - 修复版"""
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


class EmailNotifier:
    """同步邮件通知器"""
    
    def __init__(self, host: str, user: str, token: str, logger=None):
        self.host = host
        self.user = user
        self.token = token
        self.last_uid: Optional[bytes] = None
        self.mail: Optional[imaplib.IMAP4_SSL] = None
        self.logger = logger
        self.text_num = EmailConfig.DEFAULT_TEXT_NUM
        self.last_successful_check: Optional[float] = None
        self.inbox_name: str = "INBOX"


    def _log(self, message: str, level: str = 'info') -> None:
        if self.logger:
            getattr(self.logger, level)(message)

    def cleanup(self) -> None:
        if self.mail:
            try:
                self.mail.logout()
            except Exception:
                pass
            finally:
                self.mail = None

    def test_connection(self) -> bool:
        test_mail = None
        try:
            test_mail = imaplib.IMAP4_SSL(self.host, timeout=EmailConfig.CONNECTION_TIMEOUT)
            
            # 获取服务器能力
            status, capabilities = test_mail.capability()
            self._log(f"[EmailNotifier] 服务器能力: {capabilities}", 'info')
            
            # 126/163邮箱需要先发送 ID 命令（正确格式）
            try:
                test_mail.id('("name" "pyIMAP" "version" "1.0.0" "os" "python")')
            except Exception as e:
                self._log(f"[EmailNotifier] ID命令失败: {e}", 'warning')
            
            test_mail.login(self.user, self.token)
            self._log(f"[EmailNotifier] 登录成功", 'info')
            
            # 列出所有文件夹
            status, folder_list = test_mail.list()
            self._log(f"[EmailNotifier] 文件夹列表状态: {status}", 'info')
            if folder_list:
                self._log(f"[EmailNotifier] 文件夹: {folder_list[:3]}...", 'info')
            
            # 查找收件箱
            self.inbox_name = self._find_inbox_name(test_mail)
            self._log(f"[EmailNotifier] 使用收件箱: {self.inbox_name}", 'info')
            
            # 尝试选择收件箱 - 使用普通SELECT而非UID
            status, data = test_mail.select(self.inbox_name)
            self._log(f"[EmailNotifier] 选择收件箱状态: {status}, 数据: {data}", 'info')
            
            if status != 'OK':
                # 尝试其他常见收件箱名称
                for folder_name in ['INBOX', '收件箱', 'inbox', 'INBOX/Archive']:
                    if folder_name == self.inbox_name:
                        continue
                    status, data = test_mail.select(folder_name)
                    self._log(f"[EmailNotifier] 尝试 {folder_name}: {status}", 'info')
                    if status == 'OK':
                        self.inbox_name = folder_name
                        break
            
            return status == 'OK'
            
        except Exception as e:
            self._log(f"[EmailNotifier] 连接测试失败: {e}", 'error')
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
            
            # 常见的收件箱名称
            inbox_patterns = ['INBOX', 'inbox', '收件箱', 'Sent', 'Trash', 'Drafts']
            
            for folder in folder_list or []:
                if isinstance(folder, bytes):
                    folder = folder.decode('utf-8', errors='ignore')
                
                # 解析文件夹名称
                if '"' in folder:
                    parts = folder.split('"')
                    folder_name = parts[-2] if len(parts) >= 2 else parts[0]
                else:
                    folder_name = folder
                
                folder_name = folder_name.strip()
                
                # 优先返回 INBOX
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
                    # 测试连接是否活跃
                    status, _ = self.mail.status('INBOX', '(MESSAGES)')
                    if status == 'OK':
                        return
                except Exception as e:
                    self._log(f"[EmailNotifier] 现有连接测试失败: {e}", 'warning')
            
            self.cleanup()
            
            # 创建新连接
            self.mail = imaplib.IMAP4_SSL(self.host, timeout=EmailConfig.CONNECTION_TIMEOUT)
            
            # 发送客户端ID
            try:
                self.mail.id('("name" "pyIMAP" "version" "1.0.0")')
            except:
                pass
            
            # 登录
            self.mail.login(self.user, self.token)
            self._log(f"[EmailNotifier] 登录成功", 'info')
            
            # 查找收件箱名称
            self.inbox_name = self._find_inbox_name(self.mail)
            
            # 选择收件箱 - 尝试多次
            status = None
            for attempt in range(3):
                try:
                    status, data = self.mail.select(self.inbox_name)
                    self._log(f"[EmailNotifier] 选择收件箱 (尝试 {attempt+1}): {status}, 数据: {data}", 'info')
                    if status == 'OK':
                        break
                except Exception as e:
                    self._log(f"[EmailNotifier] 选择收件箱异常: {e}", 'warning')
                time.sleep(0.5)
            
            if status != 'OK':
                self._log(f"[EmailNotifier] 选择收件箱最终失败: {status}", 'error')
                self.cleanup()
                raise ConnectionError(f"无法选择收件箱: {status}")
                
        except Exception as e:
            self._log(f"[EmailNotifier] 连接失败: {e}", 'error')
            self.cleanup()
            raise

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
                subject = decoded.decode() if isinstance(decoded, bytes) else decoded
            except Exception:
                subject = msg['Subject']
        
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
                    decoded = payload.decode(part.get_content_charset() or 'utf-8')
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
                    text = payload.decode(msg.get_content_charset() or 'utf-8')
                    if msg.get_content_type() == "text/html":
                        text = self._html_to_text(text)
                    content = self._process_content(text)
            except Exception:
                pass
        
        return subject, content

    def _process_content(self, text: str) -> str:
        if not text:
            return "（无文本内容）"
        text = ' '.join(text.split())
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
            
            # 使用普通搜索而非UID搜索（更兼容）
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
            self._log(f"[EmailNotifier] 检查错误: {e}", 'error')
            self.cleanup()
            return None

    def _get_email_info(self, uid: bytes) -> Optional[Tuple[Optional[datetime], str, str]]:
        try:
            # 使用普通FETCH而非UID FETCH
            typ, msg_data = self.mail.fetch(uid, '(RFC822)')
            if typ != 'OK' or not msg_data or not msg_data[0]:
                return None

            msg = email_stdlib.message_from_bytes(msg_data[0][1])
            
            local_date = None
            date_tuple = email_stdlib.utils.parsedate_tz(msg['Date'])
            if date_tuple:
                try:
                    local_date = datetime.fromtimestamp(email_stdlib.utils.mktime_tz(date_tuple))
                except:
                    pass

            subject, content = self._get_email_content(msg)
            return local_date, subject, content
            
        except Exception as e:
            self._log(f"[EmailNotifier] 获取邮件失败: {e}", 'error')
            return None