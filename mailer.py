"""
공지 메일 자동발송 프로그램 v2
비개발자 친화 UI — 모든 설정을 프로그램 안에서 해결
"""

import os
import re
import sys
import time
import queue
import smtplib
import threading
import webbrowser
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

# ────────────────────────────────────────────────────────────────────
# 경로
# ────────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(BASE_DIR, "설정.txt")
LOG_FILE      = os.path.join(BASE_DIR, "mail_log.txt")
TEMPLATE_DIR  = os.path.join(BASE_DIR, "templates")
RECIPIENT_DIR = os.path.join(BASE_DIR, "recipients")

SMTP_PRESETS = {
    "naver":   ("smtp.naver.com",     465),
    "gmail":   ("smtp.gmail.com",     587),
    "outlook": ("smtp.office365.com", 587),
    "daum":    ("smtp.daum.net",      465),
}

SERVICE_DISPLAY = {
    "naver": "Naver", "gmail": "Gmail",
    "outlook": "Outlook", "daum": "Daum",
}

APP_PW_GUIDE = {
    "naver":   "https://nid.naver.com/user2/help/myInfoV2?m=viewSecurity",
    "gmail":   "https://myaccount.google.com/apppasswords",
    "outlook": "https://account.microsoft.com/security",
    "daum":    "https://accounts.kakao.com/login",
}

# ────────────────────────────────────────────────────────────────────
# 설정 읽기 / 저장
# ────────────────────────────────────────────────────────────────────
def load_settings():
    s = {"service": "naver", "host": "", "port": "",
         "email": "", "password": "", "sender_name": "",
         "chunk": "90", "delay": "2"}
    if not os.path.exists(SETTINGS_FILE):
        return s

    raw = {}
    with open(SETTINGS_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            raw[k.strip()] = v.strip()

    svc = raw.get("메일서비스", "").strip().lower().replace(" ", "")
    # 구버전 호환: "네이버"→"naver" 등
    svc_map = {"네이버": "naver", "naver": "naver",
               "gmail": "gmail", "outlook": "outlook",
               "office365": "outlook", "다음": "daum",
               "daum": "daum", "카카오": "daum"}
    svc = svc_map.get(svc, "직접입력" if svc else "naver")

    s["service"]     = svc if svc in SMTP_PRESETS else "직접입력"
    s["host"]        = raw.get("SMTP서버", "").strip()
    s["port"]        = raw.get("SMTP포트", "").strip()
    s["email"]       = raw.get("이메일주소", "").strip()
    s["password"]    = raw.get("앱비밀번호", "").strip()
    s["sender_name"] = raw.get("보내는사람이름", "").strip()
    s["chunk"]       = raw.get("한번에보낼최대인원", "90").strip() or "90"
    s["delay"]       = raw.get("메일간대기시간(초)", "2").strip() or "2"
    return s


def save_settings(s):
    svc = s["service"]
    if svc in SMTP_PRESETS:
        h, p = SMTP_PRESETS[svc]
    else:
        h, p = s.get("host", "").strip(), s.get("port", "").strip()

    lines = [
        "# 공지 메일 자동발송 프로그램 - 환경설정 (프로그램이 자동 저장합니다)",
        f"메일서비스 = {SERVICE_DISPLAY.get(svc, '직접입력')}",
        f"SMTP서버 = {h}",
        f"SMTP포트 = {p}",
        f"이메일주소 = {s.get('email','').strip()}",
        f"앱비밀번호 = {s.get('password','').strip()}",
        f"보내는사람이름 = {s.get('sender_name','').strip()}",
        f"한번에보낼최대인원 = {s.get('chunk','90')}",
        f"메일간대기시간(초) = {s.get('delay','2')}",
    ]
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def smtp_params(s):
    """설정 dict → (host, port, user, pw, sender_name)"""
    svc = s.get("service", "naver")
    if svc in SMTP_PRESETS and not s.get("host"):
        host, port = SMTP_PRESETS[svc]
    else:
        host = s.get("host", "").strip()
        try:
            port = int(s.get("port", "587"))
        except Exception:
            port = 587
    return host, port, s.get("email","").strip(), s.get("password","").strip(), s.get("sender_name","").strip()


# ────────────────────────────────────────────────────────────────────
# 수신자
# ────────────────────────────────────────────────────────────────────
EMAIL_RE = re.compile(r"^[^\s@\x00-\x1f]+@[^\s@\x00-\x1f]+\.[^\s@\x00-\x1f]+$")


def parse_emails(text):
    result = []
    for tok in re.split(r"[,;\n\r]+", text):
        tok = tok.strip()
        if tok and tok.lower() not in ("email", "이메일", "주소") and EMAIL_RE.match(tok):
            result.append(tok)
    return result


def load_recipient_groups():
    groups = {}
    if not os.path.isdir(RECIPIENT_DIR):
        return groups
    for fname in sorted(os.listdir(RECIPIENT_DIR)):
        if fname.endswith((".csv", ".txt")):
            name = os.path.splitext(fname)[0]
            path = os.path.join(RECIPIENT_DIR, fname)
            try:
                with open(path, encoding="utf-8", errors="ignore") as f:
                    groups[name] = parse_emails(f.read())
            except Exception:
                groups[name] = []
    return groups


def save_recipient_group(name, emails):
    os.makedirs(RECIPIENT_DIR, exist_ok=True)
    path = os.path.join(RECIPIENT_DIR, f"{name}.csv")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(emails) + "\n")


def delete_recipient_group(name):
    for ext in (".csv", ".txt"):
        p = os.path.join(RECIPIENT_DIR, f"{name}{ext}")
        if os.path.exists(p):
            os.remove(p)


# ────────────────────────────────────────────────────────────────────
# 템플릿
# ────────────────────────────────────────────────────────────────────
def list_templates():
    if not os.path.isdir(TEMPLATE_DIR):
        return []
    names = set()
    for fname in os.listdir(TEMPLATE_DIR):
        base, ext = os.path.splitext(fname)
        if ext in (".html", ".txt"):
            names.add(base)
    return sorted(names)


def get_template_col_numbers(template_name):
    """템플릿에서 사용된 {colN} 번호 목록 (정렬)"""
    for ext in (".html", ".txt"):
        path = os.path.join(TEMPLATE_DIR, f"{template_name}{ext}")
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                content = f.read()
            nums = sorted(set(int(n) for n in re.findall(r"\{col(\d+)\}", content)))
            return nums
    return []


def render_template(template_name, col_values: dict):
    """반환: (본문, content_type)"""
    for ext, ctype in [(".html", "html"), (".txt", "plain")]:
        path = os.path.join(TEMPLATE_DIR, f"{template_name}{ext}")
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                body = f.read()
            for k, v in col_values.items():
                body = body.replace(f"{{{k}}}", v)
            return body, ctype
    raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_name}")


# ────────────────────────────────────────────────────────────────────
# 메일 발송
# ────────────────────────────────────────────────────────────────────
def write_log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")
    return line


def send_mail(s, subject, body, bcc_list, content_type="plain"):
    host, port, user, pw, sender_name = smtp_params(s)
    if not host or not user or not pw:
        raise ValueError("계정 설정이 완료되지 않았습니다. [1. 계정 설정] 탭을 확인하세요.")

    msg = MIMEMultipart("alternative")
    msg["From"] = formataddr((str(Header(sender_name, "utf-8")), user)) if sender_name else user
    msg["To"] = user
    msg["Subject"] = subject
    msg.attach(MIMEText(body, content_type, "utf-8"))

    all_rcpt = [user] + bcc_list
    srv = (smtplib.SMTP_SSL(host, port, timeout=30) if port == 465
           else smtplib.SMTP(host, port, timeout=30))
    try:
        if port != 465:
            srv.starttls()
        srv.login(user, pw)
        srv.sendmail(user, all_rcpt, msg.as_string())
    finally:
        try:
            srv.quit()
        except Exception:
            pass


def send_campaign(s, subject, body, recipients, content_type,
                  test_mode, log_cb, stop_evt):
    try:
        chunk = int(s.get("chunk", "90") or "90")
    except Exception:
        chunk = 90
    try:
        delay = float(s.get("delay", "2") or "2")
    except Exception:
        delay = 2.0

    if test_mode:
        try:
            send_mail(s, f"[테스트] {subject}", body, [], content_type)
            log_cb(write_log("✅ 테스트 발송 완료 — 내 메일함을 확인하세요!"))
        except Exception as e:
            log_cb(write_log(f"❌ 테스트 발송 실패: {e}"))
        return

    if not recipients:
        log_cb(write_log("⚠️ 수신자가 없습니다. [2. 수신자 관리] 탭에서 이메일 목록을 추가하세요."))
        return

    total = 0
    for i, start in enumerate(range(0, len(recipients), chunk), 1):
        if stop_evt.is_set():
            log_cb(write_log("⏹ 발송이 중단되었습니다."))
            return
        grp = recipients[start:start + chunk]
        try:
            send_mail(s, subject, body, grp, content_type)
            total += len(grp)
            log_cb(write_log(f"✅ {i}번째 묶음 발송 완료 ({len(grp)}명)"))
        except Exception as e:
            log_cb(write_log(f"❌ {i}번째 묶음 발송 실패: {e}"))
        if start + chunk < len(recipients):
            time.sleep(delay)

    log_cb(write_log(f"🎉 발송 완료! 총 {total}명에게 발송되었습니다."))


# ────────────────────────────────────────────────────────────────────
# GUI
# ────────────────────────────────────────────────────────────────────
PAD = dict(padx=10, pady=6)


class App:
    def __init__(self, root):
        self.root = root
        root.title("공지 메일 자동발송")
        root.geometry("820x700")
        root.resizable(True, True)

        self.settings = load_settings()
        self.log_q    = queue.Queue()
        self.stop_evt = threading.Event()
        self.thread   = None
        self._groups  = {}

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        t1 = ttk.Frame(nb)
        t2 = ttk.Frame(nb)
        t3 = ttk.Frame(nb)
        nb.add(t1, text="  1. 계정 설정  ")
        nb.add(t2, text="  2. 수신자 관리  ")
        nb.add(t3, text="  3. 메일 발송  ")

        self._build_settings_tab(t1)
        self._build_recipients_tab(t2)
        self._build_send_tab(t3)

        root.after(200, self._poll)

    # ── Tab 1: 계정 설정 ──────────────────────────────────────────
    def _build_settings_tab(self, parent):
        s = self.settings
        outer = ttk.Frame(parent)
        outer.pack(fill="both", expand=True, padx=24, pady=18)

        # 서비스 선택
        ttk.Label(outer, text="메일 서비스 선택", font=("", 11, "bold")).pack(anchor="w")
        ttk.Label(outer, text="발송에 사용할 이메일 서비스를 클릭하세요.",
                  foreground="#666").pack(anchor="w", pady=(2, 10))

        svc_row = ttk.Frame(outer)
        svc_row.pack(anchor="w", pady=(0, 14))
        self._svc_var = tk.StringVar(value=s.get("service", "naver"))

        for key, label in [("naver", "  네이버  "), ("gmail", "  Gmail  "),
                            ("outlook", "  Outlook  "), ("daum", "  Daum  "),
                            ("직접입력", "  직접 입력  ")]:
            ttk.Radiobutton(svc_row, text=label, value=key,
                            variable=self._svc_var,
                            command=self._on_service_change).pack(side="left", padx=3)

        # 계정 정보 입력
        ff = ttk.LabelFrame(outer, text="계정 정보 입력")
        ff.pack(fill="x", pady=(0, 12))
        ff.columnconfigure(1, weight=1)

        self._email_var = tk.StringVar(value=s.get("email", ""))
        self._pw_var    = tk.StringVar(value=s.get("password", ""))
        self._name_var  = tk.StringVar(value=s.get("sender_name", ""))
        self._host_var  = tk.StringVar(value=s.get("host", ""))
        self._port_var  = tk.StringVar(value=s.get("port", ""))
        self._pw_visible = False

        # 이메일
        ttk.Label(ff, text="이메일 주소 *").grid(row=0, column=0, sticky="e", **PAD)
        ttk.Entry(ff, textvariable=self._email_var, width=36).grid(
            row=0, column=1, sticky="w", **PAD)
        ttk.Label(ff, text="예) myemail@naver.com",
                  foreground="#888").grid(row=0, column=2, sticky="w", padx=4)

        # 앱 비밀번호
        ttk.Label(ff, text="앱 비밀번호 *").grid(row=1, column=0, sticky="e", **PAD)
        self._pw_entry = ttk.Entry(ff, textvariable=self._pw_var, width=36, show="●")
        self._pw_entry.grid(row=1, column=1, sticky="w", **PAD)

        pw_btn_row = ttk.Frame(ff)
        pw_btn_row.grid(row=1, column=2, sticky="w", padx=4)
        ttk.Button(pw_btn_row, text="👁 보기", width=7,
                   command=self._toggle_pw).pack(side="left", padx=2)
        self._pw_guide_btn = ttk.Button(pw_btn_row, text="발급 방법 열기 →",
                                        command=self._open_pw_guide)
        self._pw_guide_btn.pack(side="left", padx=2)

        ttk.Label(ff,
                  text="⚠  일반 로그인 비밀번호가 아닌 '앱 비밀번호'를 입력하세요.  (사용설명서 3장 참고)",
                  foreground="#c00").grid(row=2, column=1, columnspan=2, sticky="w",
                                          padx=PAD["padx"], pady=(0, 4))

        # 보내는 사람 이름
        ttk.Label(ff, text="보내는 사람 이름").grid(row=3, column=0, sticky="e", **PAD)
        ttk.Entry(ff, textvariable=self._name_var, width=36).grid(
            row=3, column=1, sticky="w", **PAD)
        ttk.Label(ff, text="예) OO고객센터  (비워두면 이메일 주소로 표시)",
                  foreground="#888").grid(row=3, column=2, sticky="w", padx=4)

        # 직접입력용 SMTP 필드 (서비스 선택에 따라 표시/숨김)
        self._custom_frame = ttk.LabelFrame(ff, text="SMTP 서버 직접 입력")
        self._custom_frame.grid(row=4, column=0, columnspan=3,
                                 sticky="ew", padx=PAD["padx"], pady=6)
        ttk.Label(self._custom_frame, text="SMTP 서버").grid(
            row=0, column=0, sticky="e", padx=8, pady=5)
        ttk.Entry(self._custom_frame, textvariable=self._host_var, width=28).grid(
            row=0, column=1, sticky="w", padx=4, pady=5)
        ttk.Label(self._custom_frame, text="포트").grid(
            row=0, column=2, sticky="e", padx=8, pady=5)
        ttk.Entry(self._custom_frame, textvariable=self._port_var, width=7).grid(
            row=0, column=3, sticky="w", padx=4, pady=5)
        ttk.Label(self._custom_frame,
                  text="포트: 465 (SSL) 또는 587 (STARTTLS)",
                  foreground="#888").grid(row=0, column=4, sticky="w", padx=6)

        # 저장 / 테스트 버튼
        btn_row = ttk.Frame(outer)
        btn_row.pack(anchor="w", pady=(0, 4))
        ttk.Button(btn_row, text="💾  설정 저장", width=16,
                   command=self._save_settings).pack(side="left", padx=(0, 10))
        ttk.Button(btn_row, text="📧  테스트 메일 발송",
                   command=self._test_send).pack(side="left")

        self._settings_msg = tk.StringVar()
        ttk.Label(outer, textvariable=self._settings_msg).pack(anchor="w", pady=(4, 0))

        self._on_service_change()  # 초기 상태 반영

    def _on_service_change(self):
        svc = self._svc_var.get()
        if svc == "직접입력":
            self._custom_frame.grid()
        else:
            self._custom_frame.grid_remove()
        if hasattr(self, "_pw_guide_btn"):
            state = "normal" if svc in APP_PW_GUIDE else "disabled"
            self._pw_guide_btn.config(state=state)

    def _toggle_pw(self):
        self._pw_visible = not self._pw_visible
        self._pw_entry.config(show="" if self._pw_visible else "●")

    def _open_pw_guide(self):
        url = APP_PW_GUIDE.get(self._svc_var.get())
        if url:
            webbrowser.open(url)

    def _save_settings(self):
        svc = self._svc_var.get()
        new_s = {
            "service":     svc,
            "host":        self._host_var.get().strip(),
            "port":        self._port_var.get().strip(),
            "email":       self._email_var.get().strip(),
            "password":    self._pw_var.get().strip(),
            "sender_name": self._name_var.get().strip(),
            "chunk":       self.settings.get("chunk", "90"),
            "delay":       self.settings.get("delay", "2"),
        }
        if not new_s["email"]:
            self._settings_msg.set("❌  이메일 주소를 입력하세요.")
            return
        if not new_s["password"]:
            self._settings_msg.set("❌  앱 비밀번호를 입력하세요.")
            return
        if svc == "직접입력" and not new_s["host"]:
            self._settings_msg.set("❌  SMTP 서버 주소를 입력하세요.")
            return
        self.settings = new_s
        save_settings(new_s)
        self._settings_msg.set("✅  설정이 저장되었습니다.")

    def _test_send(self):
        self._save_settings()
        if self._settings_msg.get().startswith("❌"):
            return
        try:
            send_mail(self.settings,
                      "[테스트] 공지 메일 자동발송",
                      "테스트 메일이 정상적으로 도착했습니다!\n설정이 완료되었습니다.", [])
            messagebox.showinfo("발송 성공",
                                "✅  테스트 메일이 발송되었습니다!\n메일함을 확인해 주세요.")
        except Exception as e:
            messagebox.showerror("발송 실패", f"❌  발송 실패:\n\n{e}")

    # ── Tab 2: 수신자 관리 ────────────────────────────────────────
    def _build_recipients_tab(self, parent):
        outer = ttk.Frame(parent)
        outer.pack(fill="both", expand=True, padx=16, pady=16)

        # 왼쪽: 편집
        edit_frame = ttk.LabelFrame(outer, text="수신자 그룹 추가 / 수정")
        edit_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        ttk.Label(edit_frame,
                  text="그룹 이름 *  (예: 휴대폰고객사, VIP회원사)").pack(
            anchor="w", padx=12, pady=(10, 2))
        self._grp_name_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self._grp_name_var, width=30).pack(
            anchor="w", padx=12, pady=(0, 10), fill="x")

        ttk.Label(edit_frame,
                  text="이메일 주소 목록  (한 줄에 하나씩 입력)").pack(anchor="w", padx=12)
        self._grp_text = scrolledtext.ScrolledText(edit_frame, width=34, height=15)
        self._grp_text.pack(fill="both", expand=True, padx=12, pady=6)

        grp_btn_row = ttk.Frame(edit_frame)
        grp_btn_row.pack(anchor="w", padx=12, pady=(0, 12))
        ttk.Button(grp_btn_row, text="💾  그룹 저장",
                   command=self._save_group).pack(side="left", padx=(0, 8))
        ttk.Button(grp_btn_row, text="✖  입력 초기화",
                   command=self._clear_grp_form).pack(side="left")

        # 오른쪽: 목록
        list_frame = ttk.LabelFrame(outer, text="저장된 수신자 그룹")
        list_frame.pack(side="left", fill="both", expand=True)

        self._grp_listbox = tk.Listbox(list_frame, selectmode="single",
                                        width=24, height=16, font=("", 10))
        self._grp_listbox.pack(fill="both", expand=True, padx=10, pady=10)
        self._grp_listbox.bind("<<ListboxSelect>>", self._on_grp_select)

        list_btn_row = ttk.Frame(list_frame)
        list_btn_row.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(list_btn_row, text="✏  수정",
                   command=self._edit_group).pack(side="left", padx=(0, 6))
        ttk.Button(list_btn_row, text="🗑  삭제",
                   command=self._delete_group).pack(side="left")

        self._refresh_groups()

    def _refresh_groups(self):
        self._groups = load_recipient_groups()
        self._grp_listbox.delete(0, tk.END)
        for name, emails in self._groups.items():
            self._grp_listbox.insert(tk.END, f"  {name}  ({len(emails)}명)")
        # 발송 탭 드롭다운도 갱신
        if hasattr(self, "_rcpt_cb"):
            keys = list(self._groups.keys())
            self._rcpt_cb["values"] = ["── 그룹 선택 ──"] + keys
            if not self._rcpt_var.get() and keys:
                self._rcpt_var.set(keys[0])

    def _get_selected_grp(self):
        sel = self._grp_listbox.curselection()
        if not sel:
            return None
        raw = self._grp_listbox.get(sel[0]).strip()
        return raw.split("  (")[0].strip()

    def _on_grp_select(self, _):
        pass

    def _clear_grp_form(self):
        self._grp_name_var.set("")
        self._grp_text.delete("1.0", tk.END)

    def _edit_group(self):
        name = self._get_selected_grp()
        if not name:
            messagebox.showwarning("선택 필요", "수정할 그룹을 목록에서 선택하세요.")
            return
        self._grp_name_var.set(name)
        self._grp_text.delete("1.0", tk.END)
        self._grp_text.insert("1.0", "\n".join(self._groups.get(name, [])))

    def _save_group(self):
        name = self._grp_name_var.get().strip()
        if not name:
            messagebox.showwarning("입력 필요", "그룹 이름을 입력하세요.")
            return
        emails = parse_emails(self._grp_text.get("1.0", tk.END))
        if not emails:
            messagebox.showwarning("입력 필요",
                                   "유효한 이메일 주소가 없습니다.\n"
                                   "이메일 형식을 확인하세요. (예: user@example.com)")
            return
        save_recipient_group(name, emails)
        messagebox.showinfo("저장 완료",
                            f"'{name}' 그룹에 {len(emails)}명이 저장되었습니다.")
        self._clear_grp_form()
        self._refresh_groups()

    def _delete_group(self):
        name = self._get_selected_grp()
        if not name:
            messagebox.showwarning("선택 필요", "삭제할 그룹을 목록에서 선택하세요.")
            return
        if messagebox.askyesno("삭제 확인",
                               f"'{name}' 그룹을 삭제하시겠습니까?\n이 작업은 취소할 수 없습니다."):
            delete_recipient_group(name)
            self._refresh_groups()

    # ── Tab 3: 메일 발송 ──────────────────────────────────────────
    def _build_send_tab(self, parent):
        outer = ttk.Frame(parent)
        outer.pack(fill="both", expand=True, padx=16, pady=14)

        # ① 발송 내용 설정
        content_frame = ttk.LabelFrame(outer, text="① 발송 내용 설정")
        content_frame.pack(fill="x", pady=(0, 8))
        content_frame.columnconfigure(1, weight=1)

        ttk.Label(content_frame, text="메일 양식").grid(
            row=0, column=0, sticky="e", **PAD)
        self._tmpl_var = tk.StringVar()
        self._tmpl_cb  = ttk.Combobox(content_frame, textvariable=self._tmpl_var,
                                       width=28, state="readonly")
        self._tmpl_cb.grid(row=0, column=1, sticky="w", **PAD)
        self._tmpl_cb.bind("<<ComboboxSelected>>", self._on_template_change)
        ttk.Button(content_frame, text="↻ 목록 새로고침",
                   command=self._refresh_templates).grid(row=0, column=2, **PAD)

        ttk.Label(content_frame, text="메일 제목 *").grid(
            row=1, column=0, sticky="e", **PAD)
        self._subject_var = tk.StringVar()
        ttk.Entry(content_frame, textvariable=self._subject_var, width=52).grid(
            row=1, column=1, columnspan=2, sticky="ew", **PAD)

        # ② 입력 항목 (동적)
        self._col_outer = ttk.LabelFrame(outer, text="② 내용 입력")
        self._col_outer.pack(fill="x", pady=(0, 8))
        self._col_inner = ttk.Frame(self._col_outer)
        self._col_inner.pack(fill="x", padx=10, pady=6)
        self._col_vars: dict[str, tk.StringVar] = {}
        ttk.Label(self._col_outer,
                  text="위에서 메일 양식을 먼저 선택하면 입력 항목이 표시됩니다.",
                  foreground="#888").pack(padx=10, pady=4)

        # ③ 수신자 / 발송 옵션
        rcpt_frame = ttk.LabelFrame(outer, text="③ 수신자 및 발송 옵션")
        rcpt_frame.pack(fill="x", pady=(0, 8))
        rcpt_frame.columnconfigure(1, weight=1)

        ttk.Label(rcpt_frame, text="수신자 그룹").grid(
            row=0, column=0, sticky="e", **PAD)
        self._rcpt_var = tk.StringVar()
        self._rcpt_cb  = ttk.Combobox(rcpt_frame, textvariable=self._rcpt_var,
                                       width=28, state="readonly")
        self._rcpt_cb.grid(row=0, column=1, sticky="w", **PAD)
        ttk.Label(rcpt_frame,
                  text="수신자 그룹은 [2. 수신자 관리] 탭에서 추가합니다.",
                  foreground="#888").grid(row=0, column=2, sticky="w", padx=4)

        self._test_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            rcpt_frame,
            text="☑  테스트 모드 — 내 메일함으로만 먼저 발송 (실제 수신자에게 보내지 않음)",
            variable=self._test_var,
        ).grid(row=1, column=0, columnspan=3, sticky="w",
               padx=PAD["padx"], pady=(0, 6))

        # 발송 버튼
        send_row = ttk.Frame(outer)
        send_row.pack(anchor="w", pady=(0, 6))
        self._send_btn = ttk.Button(send_row, text="▶  발송 시작",
                                     command=self._start_send)
        self._send_btn.pack(side="left", padx=(0, 10))
        self._stop_btn = ttk.Button(send_row, text="⏹  중단",
                                     command=self._stop_send, state="disabled")
        self._stop_btn.pack(side="left")

        # 로그
        log_frame = ttk.LabelFrame(outer, text="발송 로그")
        log_frame.pack(fill="both", expand=True)
        self._log_widget = scrolledtext.ScrolledText(
            log_frame, height=8, state="disabled", font=("Consolas", 9))
        self._log_widget.pack(fill="both", expand=True, padx=6, pady=6)

        # 초기 데이터 로드
        self._refresh_templates()
        self._refresh_groups()

    def _refresh_templates(self):
        templates = list_templates()
        self._tmpl_cb["values"] = templates
        if templates:
            self._tmpl_var.set(templates[0])
            self._on_template_change()

    def _on_template_change(self, _=None):
        # col 입력 영역 초기화
        for w in self._col_inner.winfo_children():
            w.destroy()
        for w in self._col_outer.winfo_children():
            if isinstance(w, ttk.Label):
                w.destroy()
        self._col_vars.clear()

        tmpl = self._tmpl_var.get()
        if not tmpl:
            return

        col_nums = get_template_col_numbers(tmpl)
        if not col_nums:
            ttk.Label(self._col_outer,
                      text="이 양식에는 별도 입력 항목이 없습니다.",
                      foreground="#888").pack(padx=10, pady=6)
            return

        self._col_inner.columnconfigure(1, weight=1)
        for i, num in enumerate(col_nums):
            key = f"col{num}"
            var = tk.StringVar()
            self._col_vars[key] = var
            ttk.Label(self._col_inner,
                      text=f"내용 {num}  (col{num})").grid(
                row=i, column=0, sticky="e", padx=10, pady=5)
            ttk.Entry(self._col_inner, textvariable=var, width=50).grid(
                row=i, column=1, sticky="ew", padx=4, pady=5)

    def _start_send(self):
        self.settings = load_settings()

        tmpl     = self._tmpl_var.get().strip()
        subject  = self._subject_var.get().strip()
        test     = self._test_var.get()
        rcpt_key = self._rcpt_var.get().strip()

        # 유효성 검사
        if not self.settings.get("email") or not self.settings.get("password"):
            messagebox.showwarning("설정 필요",
                                   "[1. 계정 설정] 탭에서 계정 정보를 입력하고\n[설정 저장]을 눌러주세요.")
            return
        if not tmpl:
            messagebox.showwarning("선택 필요", "메일 양식을 선택해 주세요.")
            return
        if not subject:
            messagebox.showwarning("입력 필요", "메일 제목을 입력해 주세요.")
            return

        col_values = {k: v.get().strip() for k, v in self._col_vars.items()}

        try:
            body, content_type = render_template(tmpl, col_values)
        except FileNotFoundError as e:
            messagebox.showerror("파일 오류", str(e))
            return

        # 수신자 확인
        recipients = []
        if not test:
            if not rcpt_key or rcpt_key.startswith("──"):
                messagebox.showwarning("선택 필요",
                                       "테스트 모드가 꺼져 있습니다.\n"
                                       "수신자 그룹을 선택하거나\n"
                                       "테스트 모드를 켠 후 진행하세요.")
                return
            self._groups = load_recipient_groups()
            recipients = self._groups.get(rcpt_key, [])
            if not recipients:
                messagebox.showwarning("수신자 없음",
                                       f"'{rcpt_key}' 그룹에 수신자가 없습니다.\n"
                                       "[2. 수신자 관리] 탭을 확인하세요.")
                return
            if not messagebox.askyesno(
                    "발송 확인",
                    f"'{rcpt_key}' 그룹의  {len(recipients)}명에게\n"
                    f"'{subject}' 메일을 발송합니다.\n\n"
                    "계속하시겠습니까?"):
                return

        self.stop_evt.clear()
        self._send_btn.config(state="disabled")
        self._stop_btn.config(state="normal")

        self.thread = threading.Thread(
            target=send_campaign,
            kwargs=dict(
                s=self.settings, subject=subject, body=body,
                recipients=recipients, content_type=content_type,
                test_mode=test, log_cb=self.log_q.put,
                stop_evt=self.stop_evt,
            ),
            daemon=True,
        )
        self.thread.start()

    def _stop_send(self):
        self.stop_evt.set()
        self._append_log(write_log("⏹  중단 요청 전송 중..."))

    def _append_log(self, line):
        self._log_widget.config(state="normal")
        self._log_widget.insert(tk.END, line + "\n")
        self._log_widget.see(tk.END)
        self._log_widget.config(state="disabled")

    def _poll(self):
        try:
            while True:
                self._append_log(self.log_q.get_nowait())
        except queue.Empty:
            pass
        if self.thread and not self.thread.is_alive():
            self._send_btn.config(state="normal")
            self._stop_btn.config(state="disabled")
            self.thread = None
        self.root.after(200, self._poll)


# ────────────────────────────────────────────────────────────────────
def main():
    os.makedirs(TEMPLATE_DIR, exist_ok=True)
    os.makedirs(RECIPIENT_DIR, exist_ok=True)
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
