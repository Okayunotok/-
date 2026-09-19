"""
酥烤麵包機(Gradio 版)

import base64
import glob
import io
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import gradio as gr
import imageio_ffmpeg
from anthropic import Anthropic
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from dotenv import load_dotenv
from openai import OpenAI

FFMPEG_EXE = imageio_ffmpeg.get_ffmpeg_exe()

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
MASCOT_PATH = BASE_DIR / "assets" / "mascot.png"

# ---------------- 固定的社團行政資訊 ----------------
# 這些通常每次開會都一樣,直接填實際資訊,不透過 AI 從逐字稿猜測。
ORG_TITLE = "東海大學哲學系系學會　例行會議會議記錄表"
SIGNOFF_TITLES = ["哲學系系主任", "哲學系系學會會長"]
MEETING_LOCATION = ""       # 例如 "H308"
MEETING_TIME = ""           # 例如 "晚上七點至八點三十分"
MEETING_CHAIR = ""          # 例如 "楊易"
MEETING_RECORDER = ""       # 例如 "黃品鑫"

ROSTER = [
    "楊易", "王乃維", "謝禮軒", "史修一", "洪芊岫", "吳沛純", "李昱諳",
    "蔡欣妤", "呂苡希", "黃品鑫", "張毓芯", "蘇宥芸", "于子亭", "蔡禹彤",
    "李宓蜜", "歐泰佑", "張予昕", "黃可馨", "許筑茵", "周心渝", "吳卉茜",
]

# 活動企劃書:上下學期通常會辦的活動,「活動企劃書」首頁點進去後選其中一個。
# (名稱, 網址路徑) — 之後要幫哪個活動做出真正的表單,就把對應的
# with demo.route(...) 換成跟「會議記錄」一樣的完整頁面就好。
ACTIVITIES = [
    ("新生入門", "/activity-freshman"),
    ("迎新茶會", "/activity-welcome-tea"),
    ("運動會", "/activity-sports-day"),
    ("耶誕晚會", "/activity-christmas"),
    ("系烤", "/activity-bbq"),
    ("未知活動", "/activity-unknown"),
]

OPENAI_TRANSCRIBE_MODEL = "whisper-1"  # 沒有時長上限,只受 25MB 檔案大小限制
ANTHROPIC_MODEL = "claude-sonnet-5"
MAX_AUDIO_MB = 25          # OpenAI 單次請求的檔案大小上限,超過會自動切段處理,不會擋掉使用者
CHUNK_MINUTES = 18         # 自動切段時,每段大約幾分鐘
CHUNK_BITRATE = "96k"      # 切段後轉檔的音質,對語音辨識夠用,檔案也不會太大
MAX_TOTAL_MINUTES = 240    # 單一音檔的合理上限(4 小時),避免異常大檔案卡住

TEMPLATES = {
    "assoc_formal": {
        "title": "會議紀錄表 (Word)",
        "desc": "",
        "output": "docx",
    },
    "coming_soon": {
        "title": "尚未開放",
        "desc": "",
        "output": "disabled",
    },
}

CN_NUMS = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
           "十一", "十二", "十三", "十四", "十五"]

ATTENDANCE_MARKS = ["⚪", "🟢", "🔴"]  # 0=未列入 1=出席 2=缺席


# ---------------- API 用戶端 ----------------

def get_openai_client():
    key = os.environ.get("OPENAI_API_KEY")
    return OpenAI(api_key=key) if key else None


def get_anthropic_client():
    key = os.environ.get("ANTHROPIC_API_KEY")
    return Anthropic(api_key=key) if key else None


def _transcribe_single(client, filepath: str) -> str:
    """轉錄單一個檔案(必須已經在25MB以內)。"""
    with open(filepath, "rb") as f:
        resp = client.audio.transcriptions.create(
            model=OPENAI_TRANSCRIBE_MODEL,
            file=f,
            language="zh",
        )
    return resp.text


def _split_audio_with_ffmpeg(filepath: str, tmp_dir: str) -> list[str]:
    """用 ffmpeg 內建的 segment 功能把音檔切成好幾段,每段約 CHUNK_MINUTES 分鐘。
    只需要 ffmpeg 本身,不需要另外安裝 ffprobe,也不用讀取整個音檔的時長資訊。"""
    pattern = os.path.join(tmp_dir, "chunk_%03d.mp3")
    cmd = [
        FFMPEG_EXE, "-y", "-i", filepath,
        "-f", "segment",
        "-segment_time", str(CHUNK_MINUTES * 60),
        "-c:a", "libmp3lame",
        "-b:a", CHUNK_BITRATE,
        "-ac", "1",
        "-reset_timestamps", "1",
        pattern,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"音檔切段失敗(ffmpeg):{result.stderr[-500:]}")

    chunk_paths = sorted(glob.glob(os.path.join(tmp_dir, "chunk_*.mp3")))
    if not chunk_paths:
        raise RuntimeError("音檔切段後沒有產生任何檔案,請確認音檔格式是否正確。")

    total_minutes = len(chunk_paths) * CHUNK_MINUTES
    if total_minutes > MAX_TOTAL_MINUTES + CHUNK_MINUTES:
        raise ValueError(
            f"這個音檔切出了 {len(chunk_paths)} 段,長度超過單次 {MAX_TOTAL_MINUTES} 分鐘的合理上限,"
            "請先分成幾次會議或自行剪短後再上傳。"
        )
    return chunk_paths


def transcribe_audio(client, filepath: str, progress: gr.Progress | None = None) -> str:
    """轉錄音檔,依檔案大小自動決定要不要先切段。
    小檔案:跟以前一樣直接送出去。
    大檔案(超過 25MB):用 ffmpeg 自動切成每段約 CHUNK_MINUTES 分鐘,分開轉錄,
    再依原始順序把逐字稿接起來,使用者不需要自己剪音檔。"""
    size_mb = os.path.getsize(filepath) / (1024 * 1024)
    if size_mb <= MAX_AUDIO_MB:
        if progress:
            progress(0.3, desc="轉錄中…")
        text = _transcribe_single(client, filepath)
        if progress:
            progress(1.0, desc="轉錄完成")
        return text

    if progress:
        progress(0.05, desc="音檔較長,先切成幾段…")

    tmp_dir = tempfile.mkdtemp()
    texts = []
    try:
        chunk_paths = _split_audio_with_ffmpeg(filepath, tmp_dir)
        n_chunks = len(chunk_paths)
        for i, chunk_path in enumerate(chunk_paths):
            # 保險:萬一某一段音質特別高、切完還是超過上限,重新轉檔成更低的位元率。
            if os.path.getsize(chunk_path) / (1024 * 1024) > MAX_AUDIO_MB:
                lowered = chunk_path + ".low.mp3"
                subprocess.run(
                    [FFMPEG_EXE, "-y", "-i", chunk_path, "-c:a", "libmp3lame", "-b:a", "48k", "-ac", "1", lowered],
                    capture_output=True, text=True,
                )
                chunk_path = lowered
            if progress:
                progress((i + 1) / n_chunks, desc=f"轉錄第 {i + 1}/{n_chunks} 段…")
            text = _transcribe_single(client, chunk_path)
            texts.append(text)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return "\n".join(texts)


# ---------------- 結構化欄位擷取 ----------------

FIELDS_TOOL = {
    "name": "record_meeting_fields",
    "description": "提交從會議逐字稿整理出的結構化欄位,用來套進正式的社團會議記錄表 Word 檔。",
    "input_schema": {
        "type": "object",
        "properties": {
            "meeting_name": {"type": "string", "description": "會議名稱;逐字稿或補充資訊沒提到就填空字串"},
            "date": {"type": "string", "description": "會議日期;沒提到就填空字串"},
            "agenda_items": {
                "type": "array",
                "description": "逐字稿中實際討論到的議題,依討論順序列出",
                "items": {
                    "type": "object",
                    "properties": {
                        "title": {"type": "string", "description": "議題標題"},
                        "explanation": {"type": "string", "description": "背景或前情提要,沒有就填空字串"},
                        "discussion": {"type": "string", "description": "與會者的意見與討論過程,沒有就填空字串"},
                        "conclusion": {"type": "string", "description": "只有在有明確結論時才填,否則留空字串"},
                    },
                    "required": ["title", "explanation", "discussion", "conclusion"],
                },
            },
            "other_motions": {
                "type": "array", "items": {"type": "string"},
                "description": "逐字稿中額外提出、不屬於原訂議程的事項,沒有就填空陣列",
            },
            "adjournment_time": {"type": "string", "description": "散會時間;沒提到就填空字串"},
        },
        "required": ["meeting_name", "date", "agenda_items", "other_motions", "adjournment_time"],
    },
}


def build_fields_prompt(transcript, meeting_name, meeting_date):
    return (
        "你是一位專業的會議記錄整理助手,請把下面的逐字稿整理成結構化資料,"
        "之後會被套進正式的社團會議記錄表 Word 檔裡,再呼叫 record_meeting_fields 工具提交結果。\n\n"
        "規則:\n"
        "- 全部使用繁體中文\n"
        "- 逐字稿或補充資訊沒有提到的欄位,同一件事放在同一欄,不要編造內容\n"
        "- agenda_items 依逐字稿中實際討論到的議題整理,conclusion 欄位只有在有明確結論時才填,否則留空字串\n\n"
        "【補充資訊】\n"
        f"會議名稱:{meeting_name or '未提供'}\n"
        f"日期:{meeting_date or '未提供'}\n\n"
        "【逐字稿內容】\n"
        f"{transcript.strip()}"
    )


def extract_fields_json(client, prompt: str) -> dict:
    resp = client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=4000,
        tools=[FIELDS_TOOL],
        tool_choice={"type": "tool", "name": "record_meeting_fields"},
        messages=[{"role": "user", "content": prompt}],
    )
    for block in resp.content:
        if block.type == "tool_use":
            return block.input
    raise ValueError("模型沒有回傳結構化欄位,請再試一次")




# ---------------- 正式社團會議記錄表(docx 產生) ----------------

def shade_cell(cell, hex_color="D9D9D9"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_text(cell, text, bold=False, center=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text or "")
    run.bold = bold
    return cell


def add_multiline(cell, lines, first=True):
    for text, bold in lines:
        if not text:
            continue
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        run = p.add_run(text)
        run.bold = bold
    return cell


def build_formal_docx(fields: dict) -> bytes:
    doc = Document()
    doc.add_paragraph()

    table = doc.add_table(rows=0, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    col_widths = [Cm(2.6), Cm(4.6), Cm(2.6), Cm(4.6)]

    def add_row():
        row = table.add_row()
        for c, w in zip(row.cells, col_widths):
            c.width = w
        return row

    def full_width_row(text, bold=True, center=True, shade=None):
        row = add_row()
        cell = row.cells[0].merge(row.cells[1]).merge(row.cells[2]).merge(row.cells[3])
        set_text(cell, text, bold=bold, center=center)
        if shade:
            shade_cell(cell, shade)
        return cell

    def label_content_row(label, content_lines):
        row = add_row()
        set_text(row.cells[0], label, bold=True)
        content_cell = row.cells[1].merge(row.cells[2]).merge(row.cells[3])
        content_cell.text = ""
        add_multiline(content_cell, content_lines)
        return row

    full_width_row(ORG_TITLE)

    pair_rows = [
        ("會議名稱", fields.get("meeting_name", ""), "會議地點", MEETING_LOCATION),
        ("會議日期", fields.get("date", ""), "會議時間", MEETING_TIME),
        ("會議主席", MEETING_CHAIR, "會議記錄", MEETING_RECORDER),
        ("應到人數", fields.get("expected_count", ""), "實到人數", fields.get("actual_count", "")),
    ]
    for l1, v1, l2, v2 in pair_rows:
        row = add_row()
        set_text(row.cells[0], l1, bold=True)
        set_text(row.cells[1], v1)
        set_text(row.cells[2], l2, bold=True)
        set_text(row.cells[3], v2)

    attendees = "、".join(fields.get("attendees") or []) or "未提供"
    absentees = "、".join(fields.get("absentees") or []) or "無"
    label_content_row("出席人員", [(attendees, False)])
    label_content_row("缺席人員", [(absentees, False)])

    agenda_items = fields.get("agenda_items") or []

    full_width_row("會　議　議　程", shade="D9D9D9")
    agenda_lines = [
        (f"{CN_NUMS[i] if i < len(CN_NUMS) else i + 1}、{item.get('title', '')}", False)
        for i, item in enumerate(agenda_items)
    ]
    label_content_row("本次會議\n議程", agenda_lines or [("(無)", False)])

    full_width_row("會　議　記　錄　欄", shade="D9D9D9")
    for idx, item in enumerate(agenda_items):
        num = CN_NUMS[idx] if idx < len(CN_NUMS) else str(idx + 1)
        lines = []
        if item.get("explanation"):
            lines.append((f"說明:{item['explanation']}", False))
        if item.get("discussion"):
            lines.append((f"討論:{item['discussion']}", False))
        if item.get("conclusion"):
            lines.append((f"決議/結論:{item['conclusion']}", False))
        if not lines:
            lines = [("(無記錄)", False)]
        label_content_row(f"議程{num}\n{item.get('title', '')}", lines)

    other_motions = fields.get("other_motions") or []
    if other_motions:
        motion_lines = [(f"{i}. {m}", False) for i, m in enumerate(other_motions, 1)]
        label_content_row("臨時動議", motion_lines)

    if fields.get("adjournment_time"):
        label_content_row("散會", [(fields["adjournment_time"], False)])

    full_width_row("簽　　　　核", shade="D9D9D9")
    titles = SIGNOFF_TITLES
    title_row = add_row()
    left = title_row.cells[0].merge(title_row.cells[1])
    right = title_row.cells[2].merge(title_row.cells[3])
    set_text(left, titles[0] if len(titles) > 0 else "", bold=True, center=True)
    set_text(right, titles[1] if len(titles) > 1 else "", bold=True, center=True)
    sign_space_row = add_row()
    sign_space_row.cells[0].merge(sign_space_row.cells[1]).text = " "
    sign_space_row.cells[2].merge(sign_space_row.cells[3]).text = " "

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------- 畫面(Gradio) ----------------

def mascot_data_uri():
    if MASCOT_PATH.exists():
        b64 = base64.b64encode(MASCOT_PATH.read_bytes()).decode()
        return f"data:image/png;base64,{b64}"
    return None


def home_hero_html():
    uri = mascot_data_uri()
    img = (
        f'<img src="{uri}" style="width:160px;max-width:45%;margin-bottom:16px;">'
        if uri
        else ""
    )
    return f"""
    <div class="hero-section">
        {img}
        <h1 class="hero-title">哲學系學會</h1>
        <p class="hero-subtitle">文書小貓</p>
    </div>
    """


def task_header_html(title, subtitle=""):
    sub = f'<p class="task-subtitle">{subtitle}</p>' if subtitle else ""
    return f"""
    <div class="task-header">
        <h2 class="task-title">{title}</h2>
        {sub}
    </div>
    """


def progress_bar_html(percent, label):
    """自己刻的進度條,不依賴 Gradio 內建的自動進度覆蓋層
    (那個在這種自訂切換畫面的架構下不會可靠顯示出來)。"""
    percent = max(0, min(100, percent))
    return f"""
    <div class="progress-wrap">
        <div class="progress-track">
            <div class="progress-fill" style="width:{percent}%;"></div>
        </div>
        <p class="progress-label">{label}</p>
    </div>
    """


def format_choices():
    choices = []
    for key, t in TEMPLATES.items():
        label = f"{t['title']} — {t['desc']}" if t["desc"] else t["title"]
        choices.append((label, key))
    return choices


CUSTOM_CSS = """
@import url('https://fonts.googleapis.com/css2?family=LXGW+WenKai+TC&display=swap');

.gradio-container {
    background-color: #F9F4E9 !important;
}

.hero-section {
    background-color: #D8BC94;
    min-height: 42vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    text-align: center;
    border-radius: 28px;
    padding: 32px 20px;
    margin-bottom: 20px;
}
.hero-title {
    font-family: 'LXGW WenKai TC', 'Microsoft JhengHei', 'PingFang TC', sans-serif;
    color: #000000;
    font-size: 2.6rem;
    margin: 0;
    font-weight: 700;
}
.hero-subtitle { color: #2B1E14; font-size: 1rem; margin: 6px 0 0; opacity: 0.8; }

.task-header { text-align: center; margin-bottom: 12px; }
.task-title {
    font-family: 'LXGW WenKai TC', 'Microsoft JhengHei', 'PingFang TC', sans-serif;
    margin: 0;
    color: #000000;
}
.task-subtitle { color: #5B4A38; margin: 4px 0 0; }

.force-black, .force-black * { color: #000000 !important; }

.progress-wrap { padding: 12px 4px 4px; }
.progress-track {
    width: 100%;
    height: 14px;
    background-color: #E6D9BE;
    border-radius: 8px;
    overflow: hidden;
}
.progress-fill {
    height: 100%;
    background-color: #E8A33D;
    transition: width 0.4s ease;
}
.progress-label {
    text-align: center;
    color: #2B1E14;
    margin: 10px 0 0;
    font-weight: 600;
}

/* 首頁的 5 個功能卡片 */
.category-card button {
    height: 120px !important;
    font-size: 1.05rem !important;
    white-space: pre-line !important;
    border-radius: 20px !important;
    background-color: #FFFBF2 !important;
    color: #000000 !important;
    border: 2px solid #E8A33D !important;
    font-weight: 700 !important;
}
.category-card button:hover { background-color: #F3DFB8 !important; }
.category-card.card-disabled button {
    border: 2px solid #D8C7A6 !important;
    color: #5B4A38 !important;
    opacity: 0.75;
}

/* 深色底的框框:逐字稿輸入框、附加資訊裡的文字輸入框、點名按鈕。
   直接掛自己的 class,不用去猜框架內部的樣式名稱。 */
.dark-input textarea, .dark-input input {
    background-color: #2B1E14 !important;
    color: #FFFFFF !important;
}
.dark-input textarea::placeholder, .dark-input input::placeholder {
    color: #FFFFFF !important;
    opacity: 0.75;
}
.roster-btn {
    background-color: #2B1E14 !important;
    color: #FFFFFF !important;
    border: 1px solid #6B4A32 !important;
    border-radius: 16px !important;
}
.roster-btn:hover {
    background-color: #6B4A32 !important;
}
.primary-action button {
    background-color: #E8A33D !important;
    color: #000000 !important;
    font-weight: 700 !important;
}
.ghost-action button {
    background-color: transparent !important;
    border: 1px solid #6B4A32 !important;
    color: #2B1E14 !important;
}
"""

THEME = gr.themes.Soft(
    primary_hue=gr.themes.colors.orange,
    neutral_hue=gr.themes.colors.stone,
    font=[gr.themes.GoogleFont("Noto Sans TC"), "Microsoft JhengHei", "sans-serif"],
)

DISCLAIMER_TEXT = (
    "音檔會傳送到我的本機做語音辨識,逐字稿會傳送到 Anthropic 做會議紀錄整理;"
    "請留意內容是否適合送出。程式不會儲存您的音檔和逐字稿。"
)


def render_stub_page(title, back_link="/", back_label="← 回首頁"):
    """尚未開放的功能頁面,先放一個統一的佔位畫面。"""
    gr.HTML(task_header_html(title))
    gr.Markdown("這個功能還在規劃中。")
    gr.Button(back_label, link=back_link, elem_classes=["ghost-action"])


# ==================== 首頁 ====================

with gr.Blocks(title="哲學系學會 文書小貓") as demo:
    gr.HTML(home_hero_html())
    gr.Markdown("### 業務")
    with gr.Row():
        gr.Button("會議記錄", link="/meeting-minutes", elem_classes=["category-card"])
        gr.Button("活動企劃書", link="/activity-plan", elem_classes=["category-card", "card-disabled"])
        gr.Button("評鑑報告", link="/evaluation-report", elem_classes=["category-card", "card-disabled"])
        gr.Button("經費核銷", link="/expense-report", elem_classes=["category-card", "card-disabled"])
        gr.Button("尚未設計", link="/coming-soon", elem_classes=["category-card", "card-disabled"])

# ==================== 會議記錄頁 ====================
with demo.route("會議記錄", "/meeting-minutes"):
    attendance_state = gr.State({name: 0 for name in ROSTER})

    gr.HTML(task_header_html("會議記錄"))

    openai_client_check = get_openai_client()
    anthropic_client_check = get_anthropic_client()
    if not openai_client_check or not anthropic_client_check:
        missing = []
        if not openai_client_check:
            missing.append("OPENAI_API_KEY")
        if not anthropic_client_check:
            missing.append("ANTHROPIC_API_KEY")
        gr.Markdown(f"還沒偵測到:{'、'.join(missing)}。請確認 .env 或 Secrets 設定。")

    with gr.Group(visible=True) as input_group:
        gr.Markdown("#### 上傳錄音檔")
        audio_in = gr.Audio(
            sources=["upload"], type="filepath",
            label="選擇音檔(mp3 / wav / m4a / webm / ogg / flac;超過 25MB 會自動切段處理,不用自己剪)",
        )
        transcribe_btn = gr.Button("開始轉成逐字稿")
        transcribe_status = gr.Markdown("")

        gr.Markdown("#### 逐字稿內容")
        transcript_box = gr.Textbox(
            label="逐字稿內容(可直接編輯修正)", lines=8, elem_classes=["dark-input"]
        )

        gr.Markdown("#### 挑選輸出文書格式")
        format_radio = gr.Radio(
            choices=format_choices(), value="assoc_formal", label="格式",
        )

        with gr.Accordion("附加資訊(選填)", open=False):
            meeting_name_box = gr.Textbox(label="會議名稱", elem_classes=["dark-input"])
            meeting_date_box = gr.Textbox(
                label="日期", placeholder="例如:2026/09/18", elem_classes=["dark-input"]
            )
            gr.Markdown("點名:每個名字按一下依序切換 ⚪未列入 → 🟢出席 → 🔴缺席 → 回到未列入")

            person_buttons = {}
            for i in range(0, len(ROSTER), 3):
                with gr.Row():
                    for name in ROSTER[i : i + 3]:
                        person_buttons[name] = gr.Button(
                            f"⚪ {name}", elem_classes=["roster-btn"]
                        )

            attendance_summary = gr.Markdown("目前:出席 0 人・缺席 0 人・應到(合計) 0 人")

        generate_btn = gr.Button(
            "開始烘烤", variant="primary", elem_classes=["primary-action"]
        )
        generate_error = gr.Markdown("")

    with gr.Group(visible=False) as progress_group:
        gr.HTML(task_header_html("整理中…", "請容我酥烤一下"))
        progress_html = gr.HTML(progress_bar_html(0, "準備中…"))

    with gr.Group(visible=False) as result_group:
        gr.Markdown("## 出爐了!")
        gr.Markdown("已經整理成 Word 格式的會議記錄表,下載後可以直接在 Word 裡微調、列印簽核。")
        result_json = gr.JSON(label="整理出來的內容(擷取自逐字稿,下載前可以先檢查一下)")
        result_file = gr.File(label="下載 .docx")
        with gr.Row():
            back_btn = gr.Button("回上一步修改", elem_classes=["ghost-action"])
            gr.Button("← 回首頁", link="/", elem_classes=["ghost-action"])

    gr.Markdown(DISCLAIMER_TEXT, elem_classes=["force-black"])

    # ---------------- 事件邏輯 ----------------

    def do_transcribe(filepath, progress=gr.Progress()):
        client = get_openai_client()
        if not client:
            return gr.update(), "未偵測到 ,請確認 .env 或 Secrets 設定。"
        if not filepath:
            return gr.update(), "請先選擇音檔。"
        try:
            text = transcribe_audio(client, filepath, progress=progress)
            return text, "轉錄完成,已經填入下面的逐字稿欄位,可以再手動修正。"
        except Exception as e:
            return gr.update(), f"轉錄失敗:{e}"

    transcribe_btn.click(
        do_transcribe, inputs=[audio_in], outputs=[transcript_box, transcribe_status]
    )

    def make_toggle_fn(name):
        def toggle(state):
            state = dict(state)
            state[name] = (state[name] + 1) % 3
            mark = ATTENDANCE_MARKS[state[name]]
            actual = [n for n in ROSTER if state[n] == 1]
            absent = [n for n in ROSTER if state[n] == 2]
            expected = actual + absent
            summary = f"目前:出席 {len(actual)} 人・缺席 {len(absent)} 人・應到(合計) {len(expected)} 人"
            return state, gr.update(value=f"{mark} {name}"), summary

        return toggle

    for _name, _btn in person_buttons.items():
        _btn.click(
            make_toggle_fn(_name),
            inputs=[attendance_state],
            outputs=[attendance_state, _btn, attendance_summary],
        )

    def do_generate(transcript, fmt_key, meeting_name, meeting_date, state):
        # 7 個輸出對應:result_json, result_file, generate_error, progress_html,
        #              input_group, progress_group, result_group
        no_result = (gr.update(), gr.update())

        def stay_on_input(msg):
            return (*no_result, msg, gr.update(),
                    gr.update(visible=True), gr.update(visible=False), gr.update(visible=False))

        anth = get_anthropic_client()
        if not anth:
            yield stay_on_input("還沒偵測到 ANTHROPIC_API_KEY,請確認 .env 或 Secrets 設定。")
            return
        if not transcript or not transcript.strip():
            yield stay_on_input("請先完成逐字稿。")
            return
        if TEMPLATES.get(fmt_key, {}).get("output") != "docx":
            yield stay_on_input("這個格式尚未開放,請選擇「會議紀錄表 (Word)」。")
            return

        # 切到「整理中」畫面,進度條從 0 開始
        yield (*no_result, "", progress_bar_html(5, "準備中…"),
               gr.update(visible=False), gr.update(visible=True), gr.update(visible=False))

        actual_names = [n for n in ROSTER if state[n] == 1]
        absent_names = [n for n in ROSTER if state[n] == 2]
        expected_names = actual_names + absent_names

        try:
            yield (*no_result, "", progress_bar_html(25, "整理逐字稿內容…"), gr.update(), gr.update(), gr.update())
            prompt = build_fields_prompt(transcript, meeting_name, meeting_date)
            fields = extract_fields_json(anth, prompt)
            fields["expected_count"] = f"{len(expected_names)}人" if expected_names else ""
            fields["actual_count"] = f"{len(actual_names)}人" if actual_names else ""
            fields["attendees"] = actual_names
            fields["absentees"] = absent_names

            yield (*no_result, "", progress_bar_html(70, "排版成 Word 檔…"), gr.update(), gr.update(), gr.update())
            docx_bytes = build_formal_docx(fields)

            tmp_dir = tempfile.mkdtemp()
            filename = (fields.get("meeting_name") or "會議紀錄") + ".docx"
            out_path = os.path.join(tmp_dir, filename)
            with open(out_path, "wb") as f:
                f.write(docx_bytes)

            yield (*no_result, "", progress_bar_html(100, "完成!"), gr.update(), gr.update(), gr.update())

            yield (fields, out_path, "", gr.update(),
                   gr.update(visible=False), gr.update(visible=False), gr.update(visible=True))
        except Exception as e:
            yield (*no_result, f"整理失敗:{e}", gr.update(),
                   gr.update(visible=True), gr.update(visible=False), gr.update(visible=False))

    generate_btn.click(
        do_generate,
        inputs=[transcript_box, format_radio, meeting_name_box, meeting_date_box, attendance_state],
        outputs=[result_json, result_file, generate_error, progress_html, input_group, progress_group, result_group],
    )

    back_btn.click(
        lambda: (gr.update(visible=True), gr.update(visible=False), gr.update(visible=False)),
        inputs=None,
        outputs=[input_group, progress_group, result_group],
    )

# ==================== 其他功能(先放佔位頁) ====================

with demo.route("活動企劃書", "/activity-plan"):
    gr.HTML(task_header_html("活動企劃書"))
    gr.Markdown("### 上學期", elem_classes=["force-black"])
    with gr.Row():
        for act_name, act_path in ACTIVITIES[:4]:
            gr.Button(act_name, link=act_path, elem_classes=["category-card"])
    gr.Markdown("### 下學期", elem_classes=["force-black"])
    with gr.Row():
        for act_name, act_path in ACTIVITIES[4:]:
            gr.Button(act_name, link=act_path, elem_classes=["category-card"])
    gr.Button("← 回首頁", link="/", elem_classes=["ghost-action"])

for _act_name, _act_path in ACTIVITIES:
    with demo.route(_act_name, _act_path, show_in_navbar=False):
        render_stub_page(_act_name, back_link="/activity-plan", back_label="← 回活動企劃書")

with demo.route("評鑑報告", "/evaluation-report"):
    render_stub_page("評鑑報告")

with demo.route("經費核銷", "/expense-report"):
    render_stub_page("經費核銷")

with demo.route("尚未設計", "/coming-soon"):
    render_stub_page("尚未設計")


# 直接寫進 <head> 的內嵌樣式,瀏覽器解析 HTML 時就會馬上套用,
# 不用等外部樣式表載入完成——這樣換頁(整頁重新整理)那一瞬間白色畫面的時間會縮短。
# 注意:這只能縮短,沒辦法完全消除——換頁本身是瀏覽器在做整頁重新整理,
# 從「按下連結」到「這個 HTML 送達瀏覽器」之間那一小段網路傳輸的空白,
# 是瀏覽器自己的行為,不是任何程式碼能控制的範圍。
HEAD_HTML = """
<style>
  html, body, #root, .gradio-container, gradio-app {
    background-color: #F9F4E9 !important;
    margin: 0;
  }
</style>
"""

if __name__ == "__main__":
    # 本機測試:維持 127.0.0.1:7860,跟之前一樣。
    # 部署到 Render 之類的平台時,平台會透過 PORT 環境變數指定要用哪個埠,
    # 且必須綁定 0.0.0.0(不能只監聽 127.0.0.1)才能讓外部連線進來。
    port = int(os.environ.get("PORT", 7860))
    is_deployed = "PORT" in os.environ
    demo.launch(
        theme=THEME, css=CUSTOM_CSS, head=HEAD_HTML,
        server_name="0.0.0.0" if is_deployed else "127.0.0.1",
        server_port=port,
    )
