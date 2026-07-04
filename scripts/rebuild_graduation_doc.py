from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
DOC_PATH = ROOT / "JARVIS_Documentation.docx"
FIGURE_DIR = ROOT / "docs" / "generated_figures"


def _font(size=24, bold=False):
    candidates = [
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/calibrib.ttf" if bold else "C:/Windows/Fonts/calibri.ttf",
    ]
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size)
        except Exception:
            pass
    return ImageFont.load_default()


def _draw_centered(draw, box, text, font, fill=(20, 30, 40)):
    lines = []
    for raw in text.split("\n"):
        words = raw.split()
        line = ""
        for word in words:
            trial = (line + " " + word).strip()
            if draw.textbbox((0, 0), trial, font=font)[2] - draw.textbbox((0, 0), trial, font=font)[0] > (box[2] - box[0] - 22):
                if line:
                    lines.append(line)
                line = word
            else:
                line = trial
        if line:
            lines.append(line)
    line_h = font.size + 6
    total_h = line_h * len(lines)
    y = box[1] + ((box[3] - box[1] - total_h) / 2)
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        x = box[0] + ((box[2] - box[0] - (bbox[2] - bbox[0])) / 2)
        draw.text((x, y), line, font=font, fill=fill)
        y += line_h


def _arrow(draw, start, end, fill=(40, 60, 80)):
    draw.line([start, end], fill=fill, width=4)
    ex, ey = end
    sx, sy = start
    if ex >= sx:
        head = [(ex, ey), (ex - 14, ey - 8), (ex - 14, ey + 8)]
    else:
        head = [(ex, ey), (ex + 14, ey - 8), (ex + 14, ey + 8)]
    draw.polygon(head, fill=fill)


def _box(draw, box, text, color):
    draw.rounded_rectangle(box, radius=14, fill=color, outline=(45, 55, 65), width=2)
    _draw_centered(draw, box, text, _font(22, bold=True))


def make_figure(filename, title, boxes):
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    path = FIGURE_DIR / filename
    img = Image.new("RGB", (1400, 520), (250, 252, 255))
    draw = ImageDraw.Draw(img)
    draw.text((48, 32), title, font=_font(34, bold=True), fill=(12, 24, 38))
    y = 190
    w = 190
    gap = 40
    x = 52
    palette = [(220, 238, 255), (226, 246, 232), (255, 240, 218), (239, 230, 255), (234, 244, 244), (255, 229, 231)]
    centers = []
    for idx, label in enumerate(boxes):
        box = (x, y, x + w, y + 112)
        _box(draw, box, label, palette[idx % len(palette)])
        centers.append((x + w, y + 56, x, y + 56))
        x += w + gap
    for idx in range(len(centers) - 1):
        _arrow(draw, (centers[idx][0] + 8, centers[idx][1]), (centers[idx + 1][2] - 8, centers[idx + 1][3]))
    img.save(path)
    return path


def add_figure(doc, number, caption, filename, title, boxes):
    path = make_figure(filename, title, boxes)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=Inches(6.6))
    para(doc, f"{number} - {caption}", style="Caption")


def set_margins(section):
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.85)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)


def add_page_number(section):
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def para(doc, text="", style=None, align=None, bold=False, size=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if text:
        run = p.add_run(text)
        run.bold = bold
        if size:
            run.font.size = Pt(size)
    return p


def bullet(doc, text):
    para(doc, text, style="List Bullet")


def numbered(doc, text):
    para(doc, text, style="List Number")


def add_table(doc, title, headers, rows):
    para(doc, title, style="Caption")
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for idx, header in enumerate(headers):
        table.rows[0].cells[idx].text = header
    for row in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            cells[idx].text = str(value)
    para(doc)


def extract_acknowledgement():
    old = Document(str(DOC_PATH))
    ack = []
    collecting = False
    for p in old.paragraphs:
        text = " ".join(p.text.split())
        if text == "Acknowledgement":
            collecting = True
            continue
        if collecting and text in {"Table of Contents", "Chapter 1: Introduction"}:
            break
        if collecting and text:
            ack.append(text)
    return ack or [
        "First and foremost, praises and thanks to Allah, the Almighty, for His blessings throughout this project.",
        "We would like to express our deepest gratitude to our project supervisor, Dr. Mostafa El Sayed, for his guidance and support.",
        "We also thank every team member for their cooperation, patience, and commitment during the development of JARVIS.",
    ]


def add_preliminaries(doc, acknowledgement):
    para(doc, "JARVIS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=28)
    para(
        doc,
        "A Local Bilingual Voice Assistant for Windows Desktop Automation",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        bold=True,
        size=16,
    )
    para(doc)
    para(doc, "Graduation Project Documentation Book", align=WD_ALIGN_PARAGRAPH.CENTER, size=14)
    para(doc)
    para(doc, "Submitted by:", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    para(doc, "[Team member names to be completed]", align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, "Supervised by:", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    para(doc, "Dr. Mostafa El Sayed", align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, "Faculty / Department: [to be completed]", align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, "Academic Year: [to be completed]", align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, "Submission Date: [to be completed]", align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    para(doc, "Acknowledgement", style="Heading 1")
    for item in acknowledgement:
        para(doc, item)
    doc.add_page_break()

    para(doc, "Abstract", style="Heading 1")
    para(
        doc,
        "JARVIS is a local-first bilingual voice assistant designed for Windows desktop automation. "
        "The project combines wake-word detection, voice activity detection, speech-to-text transcription, "
        "bilingual intent routing, local large-language-model fallback, text-to-speech output, and Windows automation "
        "into a single Python-based runtime.",
    )
    para(
        doc,
        "The implemented system begins with a unified OpenWakeWord-compatible ONNX wake model, records speech through "
        "a streaming microphone pipeline, applies VAD using Silero ONNX where available with an energy fallback, "
        "transcribes English and Egyptian Arabic through Faster-Whisper or an optional ElevenLabs hybrid backend, "
        "routes commands through deterministic and semantic NLU components, and executes supported actions through "
        "dedicated OS-control modules. Open-ended questions are routed to an Ollama-backed local LLM unless the "
        "configured backend is changed.",
    )
    para(
        doc,
        "This documentation avoids treating planned or optional features as guaranteed behavior. Technical claims are "
        "based on the current repository, configuration files, tests, and the existing project document. Unsupported "
        "benchmark values from the earlier draft were removed or converted into evaluation placeholders. The current "
        "automated test suite was verified during this documentation pass with 41 passing tests and one third-party "
        "deprecation warning.",
    )
    doc.add_page_break()

    para(doc, "Table of Contents", style="Heading 1")
    toc_lines = [
        "Acknowledgement",
        "Abstract",
        "List of Figures",
        "List of Tables",
        "List of Symbols / Abbreviations",
        "Chapter 1: Introduction",
        "Chapter 2: Literature Review",
        "Chapter 3: Methodology",
        "Chapter 4: Implementation / Analysis",
        "Chapter 5: Conclusions & Future Work",
        "References",
        "Appendices",
    ]
    for line in toc_lines:
        para(doc, line)
    doc.add_page_break()

    para(doc, "List of Figures", style="Heading 1")
    add_table(
        doc,
        "Table LF.1 - List of Figures",
        ["Figure No.", "Caption", "Status"],
        [
            ["Figure 1.1", "Conceptual overview of JARVIS voice interaction", "Textual placeholder"],
            ["Figure 3.1", "Methodology flow from wake word to response", "Textual placeholder"],
            ["Figure 4.1", "Repository package structure", "Textual placeholder"],
            ["Figure 4.2", "Runtime command-processing pipeline", "Textual placeholder"],
        ],
    )
    para(doc, "List of Tables", style="Heading 1")
    add_table(
        doc,
        "Table LT.1 - List of Tables",
        ["Table No.", "Caption"],
        [
            ["Table 2.1", "Comparison of existing systems and JARVIS"],
            ["Table 3.1", "Development methodology stages"],
            ["Table 4.1", "Main repository folders and responsibilities"],
            ["Table 4.2", "Core dependencies verified from requirements.txt"],
            ["Table 4.3", "Selected environment variables verified from core/config.py"],
            ["Table 4.4", "LLM model tier selection verified from core/hardware_detect.py"],
            ["Table 4.5", "Automated test coverage verified from tests/"],
            ["Table A.1", "Source-code listing reference"],
            ["Table B.1", "Installation guide"],
            ["Table F.1", "Contribution table placeholder"],
        ],
    )
    para(doc, "List of Symbols / Abbreviations", style="Heading 1")
    add_table(
        doc,
        "Table LS.1 - Abbreviations",
        ["Abbreviation", "Meaning"],
        [
            ["AI", "Artificial Intelligence"],
            ["ASR / STT", "Automatic Speech Recognition / Speech-to-Text"],
            ["GPU", "Graphics Processing Unit"],
            ["KB", "Knowledge Base"],
            ["LLM", "Large Language Model"],
            ["NLU", "Natural Language Understanding"],
            ["ONNX", "Open Neural Network Exchange"],
            ["TTS", "Text-to-Speech"],
            ["VAD", "Voice Activity Detection"],
            ["WAL", "Write-Ahead Logging"],
        ],
    )
    doc.add_page_break()


def add_chapter_1(doc):
    para(doc, "Chapter 1: Introduction", style="Heading 1")
    sections = [
        (
            "1.1 General Introduction",
            "Voice-controlled computing is common in mobile and smart-home environments, yet desktop automation remains "
            "less accessible for users who need hands-free control of files, applications, settings, and everyday "
            "Windows workflows. JARVIS addresses this gap by implementing a local bilingual voice assistant focused on "
            "Windows desktop automation in English and Egyptian Arabic.",
        ),
        (
            "1.2 Problem Statement",
            "The problem addressed by this project is the lack of a locally deployable bilingual assistant that combines "
            "voice interaction, Egyptian Arabic support, and practical Windows desktop automation. Existing assistants "
            "often emphasize cloud services, mobile or smart-home workflows, limited desktop command surfaces, or "
            "language coverage that does not directly target Egyptian colloquial Arabic.",
        ),
        (
            "1.3 Project Objectives",
            "The primary objective is to design, implement, and document a Windows-focused voice assistant that can "
            "listen, understand, route, execute, and respond through a local-first architecture.",
        ),
    ]
    for heading, body in sections:
        para(doc, heading, style="Heading 2")
        para(doc, body)
    for obj in [
        "Support English and Egyptian Arabic input through unified wake-word and bilingual processing.",
        "Implement microphone capture, VAD, STT, command routing, LLM fallback, TTS, and Windows automation.",
        "Store runtime context through the implemented session-memory and SQLite memory layers.",
        "Handle risky actions through policy and confirmation components.",
        "Document limitations honestly without presenting planned work as implemented.",
    ]:
        numbered(doc, obj)
    para(doc, "1.4 Scope of Work", style="Heading 2")
    para(
        doc,
        "The current implementation targets Windows desktop use. The repository contains Python backend packages for "
        "audio, core orchestration, NLP, LLM integration, OS control, tools, utilities, and tests. It also contains a "
        "desktop/ folder with a Tauri/React UI project and WebSocket bridge integration points.",
    )
    para(
        doc,
        "Out of scope are unsupported operating systems, guaranteed offline behavior for optional cloud-backed services, "
        "fully validated large-scale accuracy benchmarks, automatic sending of email without user review, and production "
        "claims not supported by repository tests or measured evaluation data.",
    )
    para(doc, "1.5 Motivation", style="Heading 2")
    para(
        doc,
        "The motivation for JARVIS is practical accessibility: a user should be able to control common Windows tasks "
        "through natural speech without depending on a commercial cloud assistant or GPU-heavy setup. The bilingual focus "
        "matters because Egyptian Arabic speakers frequently code-switch between Arabic and English, especially when naming "
        "applications, files, and technical actions.",
    )
    para(doc, "1.6 Report Organization", style="Heading 2")
    para(
        doc,
        "Chapter 2 reviews related technologies and systems. Chapter 3 explains the methodology used to design the assistant. "
        "Chapter 4 analyzes the actual implementation, repository structure, dependencies, runtime pipeline, testing, and "
        "evaluation status. Chapter 5 presents conclusions, achieved objectives, limitations, and future improvements.",
    )
    add_figure(
        doc,
        "Figure 1.1",
        "Conceptual overview of JARVIS voice interaction",
        "figure_1_1_overview.png",
        "JARVIS Voice Interaction Overview",
        ["Wake Word", "Record Speech", "STT", "Intent Routing", "Action or LLM", "TTS Response"],
    )
    doc.add_page_break()


def add_chapter_2(doc):
    para(doc, "Chapter 2: Literature Review", style="Heading 1")
    sections = [
        ("2.1 Previous Work", "Voice assistants combine speech recognition, language understanding, action execution, and speech synthesis. JARVIS builds on this architecture but adapts it for local Windows desktop automation."),
        ("2.2 Related Systems", "Related systems include commercial assistants, open-source wake-word engines, local STT frameworks, VAD models, and local LLM runners. JARVIS integrates these categories rather than introducing a new model architecture from scratch."),
        ("2.3 Existing Voice Assistants", "Commercial assistants demonstrate mature voice interaction, but their design assumptions commonly differ from this project: they emphasize cloud services, mobile or smart-home workflows, and controlled integration surfaces."),
        ("2.4 Local AI Assistants", "Local AI assistants are more practical because optimized inference tools such as Ollama and CTranslate2-based STT allow capable models to run on consumer hardware."),
        ("2.5 Speech Recognition Technologies", "The project uses Faster-Whisper for local transcription and optionally uses ElevenLabs STT through the hybrid backend configured in audio/stt.py and core/config.py."),
        ("2.6 Wake-Word Detection", "Wake-word detection is handled by audio/wake_word.py through a unified OpenWakeWord-compatible ONNX model configured by JARVIS_WAKE_WORD_UNIFIED_ONNX_PATH."),
        ("2.7 Voice Activity Detection", "JARVIS uses audio/vad.py, which loads a Silero ONNX model when possible and falls back to energy gating when ONNX runtime or model loading is unavailable."),
        ("2.8 Natural Language Understanding", "NLU is implemented as a cascade rather than a single model call. The repository includes command_parser, code_switch_router, semantic_router, keyword_engine, fuzzy_matcher, intent_confidence, and nlu modules."),
        ("2.9 Arabic and Egyptian Arabic NLP Challenges", "Egyptian Arabic introduces dialectal vocabulary, spelling variation, mixed Arabic/Latin scripts, and frequent code-switching with English application names."),
        ("2.10 Research Gap", "The identified gap is a local, Windows-focused, bilingual assistant that treats Egyptian Arabic as a first-class spoken interaction mode while supporting everyday desktop automation."),
        ("2.11 Comparison with Existing Systems", "Table 2.1 summarizes the comparison at a system level. Cells describe general architectural fit rather than exhaustive product capabilities."),
    ]
    for heading, body in sections:
        para(doc, heading, style="Heading 2")
        para(doc, body)
    add_table(
        doc,
        "Table 2.1 - Comparison of Existing Systems and JARVIS",
        ["System Type", "Primary Strength", "Limitation for This Project", "JARVIS Position"],
        [
            ["Commercial assistants", "Mature voice interaction", "Cloud-centered and limited desktop control", "Local-first Windows automation"],
            ["Wake-word tools", "Efficient local activation", "Activation only", "Used as one subsystem"],
            ["Local STT engines", "Offline transcription potential", "No desktop action layer", "Integrated into full runtime"],
            ["Local LLM chat tools", "Flexible conversation", "Risky or slow if used for every OS action", "Fallback after routing"],
        ],
    )
    doc.add_page_break()


def add_chapter_3(doc):
    para(doc, "Chapter 3: Methodology", style="Heading 1")
    sections = [
        ("3.1 Development Methodology", "The project follows an incremental systems-integration methodology. Subsystems are developed as separate packages, then connected through the orchestrator and command router."),
        ("3.2 Requirements Analysis", "Requirements were derived from hands-free Windows control, bilingual interaction, local execution where possible, safe handling of risky actions, and graceful degradation."),
        ("3.3 System Design Approach", "Responsibilities are separated by package: audio handles speech I/O, core handles orchestration and policy, nlp handles classification support, llm handles model calls, os_control handles Windows actions, and tools handles live data."),
        ("3.4 Architecture Methodology", "The architecture uses explicit pipeline boundaries: wake detection, recording, transcription, language detection, parsing/routing, action execution or LLM response, memory update, and speech synthesis."),
        ("3.5 Audio Pipeline Methodology", "Audio capture is configured around a 16 kHz sample rate in core/config.py. Streaming capture and partial transcription allow early command execution and live-data prefetch in selected cases."),
        ("3.6 Wake-Word Detection Methodology", "The wake-word method loads one configured ONNX wake model, applies threshold and confirm-frame logic, observes a cooldown interval, and records recent activation audio for adaptive wake behavior."),
        ("3.7 VAD Methodology", "The VAD method first checks energy to avoid unnecessary model calls. When available, Silero ONNX inference is used on fixed windows; otherwise energy gating is used as a fallback."),
        ("3.8 STT Methodology", "The STT method supports local Faster-Whisper and hybrid ElevenLabs backends. The code locks transcription to English or Arabic when configured, validates transcripts by script composition, and can retry with the opposite language."),
        ("3.9 Bilingual Processing Methodology", "Bilingual handling combines language hints, script detection, romanized Arabic checks, code-switch routing, Arabic normalization, and response-language selection."),
        ("3.10 NLU and Intent Routing Methodology", "Routing begins with deterministic parsing and can proceed through code-switch handling, semantic classification, keyword/fuzzy logic, slot filling, clarification, dispatch, and fallback response generation."),
        ("3.11 LLM Fallback Methodology", "The LLM is used for open-ended questions, low-confidence inputs, and conversational cases. Ollama is the default local backend; Claude support exists as an optional configured backend."),
        ("3.12 Memory and Persistence Methodology", "Session memory stores recent turns, language preferences, pending clarifications, and context slots. SQLiteMemoryStore uses turns and slots tables with WAL mode; VectorMemoryStore optionally uses ChromaDB."),
        ("3.13 Risk Handling and Confirmation Methodology", "Risky operations are routed through confirmation and policy components, including spoken PIN confirmation for sensitive commands and policy profiles."),
        ("3.14 Error Handling and Fallback Strategy", "Optional import failures and runtime failures are handled through degraded features and fallback paths where possible."),
        ("3.15 Privacy and Offline-First Methodology", "The design keeps wake detection, VAD, local STT, command routing, local LLM, memory, and core OS actions on the user machine when configured accordingly. Optional services require network access."),
    ]
    for heading, body in sections:
        para(doc, heading, style="Heading 2")
        para(doc, body)
    add_figure(
        doc,
        "Figure 3.1",
        "Methodology flow from wake word to spoken response",
        "figure_3_1_methodology.png",
        "Methodology Flow",
        ["Requirements", "Audio Method", "NLU Method", "Risk Controls", "Persistence", "Evaluation"],
    )
    add_table(
        doc,
        "Table 3.1 - Development Methodology Stages",
        ["Stage", "Repository Evidence", "Output"],
        [
            ["Subsystem separation", "audio/, core/, nlp/, llm/, os_control/, tools/", "Maintainable package structure"],
            ["Runtime orchestration", "main.py and core/orchestrator.py", "Wake-to-response loop"],
            ["Command routing", "core/command_router.py and core/command_parser.py", "Intent selection and dispatch"],
            ["Persistence", "core/session_memory.py and core/memory_store.py", "Recent turns, slots, and optional semantic memory"],
            ["Verification", "tests/ and pytest run", "41 passing automated tests in this pass"],
        ],
    )
    doc.add_page_break()


def add_chapter_4(doc):
    para(doc, "Chapter 4: Implementation / Analysis", style="Heading 1")
    para(doc, "4.1 Actual Codebase Structure", style="Heading 2")
    para(doc, "The implementation is organized as a multi-package repository. The main executable entry point is main.py. Runtime configuration is centralized in core/config.py and environment variables are provided through .env and .env.example.")
    add_table(
        doc,
        "Table 4.1 - Main Repository Folders and Responsibilities",
        ["Folder/File", "Responsibility"],
        [
            ["main.py", "Starts optional tray and UI bridge, then calls core.orchestrator.run()."],
            ["audio/", "Wake-word listening, microphone capture, VAD, STT, TTS, cues, and wake enrollment."],
            ["core/", "Orchestration, routing, configuration, memory, metrics, policies, persona, diagnostics, and handlers."],
            ["nlp/", "Semantic routing, code-switch routing, keyword/fuzzy matching, NLU helpers, and entity types."],
            ["llm/", "Ollama client, prompt builders, sentence buffering, tool calling, and prompt templates."],
            ["os_control/", "Windows automation adapters for apps, files, system controls, reminders, timers, clipboard, settings, email, calendar, and policy."],
            ["tools/", "Weather, web search, live-data aggregation, and calculator tools."],
            ["ui/ and desktop/", "Python tray/bridge and optional Tauri/React desktop interface."],
            ["tests/", "Automated unit tests for routing guards, sentence buffering, TTS prosody, and voice normalization."],
        ],
    )
    for heading, body in [
        ("4.2 Main Package Structure", "The package structure follows the runtime pipeline. audio modules prepare input and output; core modules control decisions and state; nlp modules provide classification support; llm modules manage fallback responses; os_control modules isolate Windows side effects; tools modules provide live-data utilities."),
        ("4.3 Main Application Workflow", "At startup, main.py parses --demo-mode, tries to start the tray and bridge, and calls run() in core/orchestrator.py. The orchestrator initializes services, configures STT, prewarms components, optionally starts background tasks, speaks a greeting, and enters the wake/listen/process loop."),
        ("4.4 Folder-by-Folder Implementation Explanation", "The folders listed in Table 4.1 map directly to runtime responsibilities, allowing the report to explain each subsystem without mixing audio capture, automation, and prompt construction."),
    ]:
        para(doc, heading, style="Heading 2")
        para(doc, body)
    add_figure(
        doc,
        "Figure 4.1",
        "Repository package structure",
        "figure_4_1_packages.png",
        "Repository Package Structure",
        ["audio", "core", "nlp", "llm", "os_control", "tools/ui"],
    )
    para(doc, "4.5 File Responsibilities", style="Heading 2")
    for item in [
        "core/orchestrator.py: coordinates startup, wake-word listening, recording, transcription, early execution, live-data prefetch, response playback, and shutdown cleanup.",
        "core/command_router.py: routes recognized text to deterministic commands, clarification flow, action handlers, tool calls, or LLM fallback.",
        "core/config.py: reads environment variables and exposes runtime defaults.",
        "audio/wake_word.py: loads and runs the unified wake-word ONNX model through openWakeWord.",
        "audio/vad.py: provides Silero ONNX VAD with energy fallback.",
        "audio/stt.py: implements Faster-Whisper and hybrid ElevenLabs STT paths.",
        "core/memory_store.py: implements SQLite and optional ChromaDB-backed vector memory stores.",
        "core/hardware_detect.py: recommends Whisper runtime settings and Qwen3 LLM tiers.",
    ]:
        bullet(doc, item)
    para(doc, "4.6 Dependencies", style="Heading 2")
    add_table(
        doc,
        "Table 4.2 - Core Dependencies Verified from requirements.txt",
        ["Dependency", "Role in JARVIS"],
        [
            ["python-dotenv", "Loads .env configuration."],
            ["numpy, sounddevice, onnxruntime", "Audio processing and ONNX inference."],
            ["faster-whisper, openwakeword", "Local STT and wake-word runtime."],
            ["edge-tts, soundfile", "Text-to-speech synthesis and decoding."],
            ["httpx, psutil, rapidfuzz", "HTTP calls, hardware info, and fuzzy matching."],
            ["chromadb, sentence-transformers", "Optional vector memory and semantic routing."],
            ["duckduckgo-search / ddgs", "Optional web search."],
            ["pywin32, pycaw, wmi, pyperclip, screen-brightness-control", "Optional Windows integrations."],
            ["fastapi, uvicorn, websockets", "Optional desktop UI bridge."],
        ],
    )
    para(doc, "4.7 Models Used", style="Heading 2")
    para(doc, "Configured models include the unified wake-word ONNX model, Silero VAD ONNX, Faster-Whisper models selected automatically or by JARVIS_WHISPER_MODEL, optional sentence-transformer models, and Qwen3 models served by Ollama.")
    add_table(
        doc,
        "Table 4.4 - LLM Model Tier Selection Verified from core/hardware_detect.py",
        ["RAM", "GPU Required", "Tier", "Model", "num_ctx", "lightweight_ctx"],
        [
            ["16 GB+", "Yes", "high", "qwen3:8b", "8192", "4096"],
            ["12 GB+", "No", "medium", "qwen3:4b", "4096", "2048"],
            ["8 GB+", "Yes", "medium", "qwen3:4b", "4096", "2048"],
            ["8 GB+", "No", "low", "qwen3:1.7b", "2048", "1024"],
            ["Below 8 GB", "No", "minimal", "qwen3:0.6b", "1024", "512"],
        ],
    )
    para(doc, "4.8 Configuration Files", style="Heading 2")
    para(doc, "Configuration is primarily handled by .env, .env.example, requirements*.txt, desktop/package.json, desktop/src-tauri/Cargo.toml, and runtime databases such as jarvis_memory.db, jarvis_state.db, and jarvis_index.db.")
    para(doc, "4.9 Environment Variables", style="Heading 2")
    add_table(
        doc,
        "Table 4.3 - Selected Environment Variables Verified from core/config.py",
        ["Variable", "Default / Behavior", "Purpose"],
        [
            ["JARVIS_UI_BRIDGE_ENABLED", "true", "Enables Python-to-desktop UI bridge."],
            ["JARVIS_MAX_RECORD_DURATION", "8.0 bounded to 3-20 seconds", "Maximum utterance recording time."],
            ["JARVIS_VAD_BACKEND", "silero", "Selects Silero or energy VAD behavior."],
            ["JARVIS_WAKE_WORD_UNIFIED_ONNX_PATH", "models/jarvis_unified/jarvis_unified.onnx", "Wake-word model path."],
            ["JARVIS_STT_BACKEND", "hybrid_elevenlabs", "Selects hybrid or local STT backend."],
            ["JARVIS_ELEVENLABS_MODEL", "scribe_v2", "ElevenLabs STT model name."],
            ["JARVIS_LLM_BACKEND", "ollama", "Selects Ollama or Claude backend."],
            ["JARVIS_MEMORY_BACKEND", "sqlite", "Selects SQLite or JSON memory backend."],
            ["JARVIS_SENSITIVE_CONFIRM_MODE", "pin", "Controls sensitive command confirmation."],
        ],
    )
    for heading, body in [
        ("4.10 Runtime Pipeline", "The runtime pipeline is: wake detection -> VAD-gated recording -> STT transcription -> language detection/validation -> command parsing and routing -> optional clarification or confirmation -> action execution or LLM/tool response -> response shaping -> TTS playback -> memory and metrics update."),
        ("4.11 Architecture", "The architecture is modular and event driven. The orchestrator owns lifecycle and concurrency, while command_router.py owns interpretation and dispatch."),
        ("4.12 Memory Scheme", "SQLiteMemoryStore creates turns and slots tables, imports legacy jarvis_memory.json once, and can export to JSON. VectorMemoryStore optionally uses ChromaDB and all-MiniLM-L6-v2 embeddings."),
        ("4.13 Command-Processing Pipeline", "The command-processing path includes parser fast paths, code-switch routing, semantic routing, keyword/fuzzy support, entity enrichment, missing-slot questions, clarification handling, confirmation, dispatch, and LLM fallback."),
    ]:
        para(doc, heading, style="Heading 2")
        para(doc, body)
    add_figure(
        doc,
        "Figure 4.2",
        "Runtime command-processing pipeline",
        "figure_4_2_pipeline.png",
        "Runtime Command Pipeline",
        ["Parser", "Code Switch", "Semantic", "Clarify", "Dispatch", "Memory/Metrics"],
    )
    para(doc, "4.14 Behavior in Different Situations", style="Heading 2")
    for item in [
        "If the wake-word or audio stack is unavailable, the orchestrator prints an error and enters a text fallback loop.",
        "If no speech is detected, temporary audio is removed and the dialogue state returns to idle where appropriate.",
        "If a partial transcript clearly matches a safe direct command twice and passes the confidence gate, early execution can run.",
        "If a command is ambiguous or missing a required slot, the assistant stores a pending clarification and asks a follow-up question.",
        "If a sensitive action requires confirmation, the confirmation manager handles the spoken PIN flow before execution.",
    ]:
        bullet(doc, item)
    para(doc, "4.15 Testing and Evaluation", style="Heading 2")
    para(doc, "Automated tests were verified in this documentation pass using python -m pytest -q. The result was 41 passing tests with one third-party deprecation warning. These tests cover pure logic and text-processing behavior; they do not constitute full hardware, STT, wake-word, or end-to-end voice evaluation.")
    add_table(
        doc,
        "Table 4.5 - Automated Test Coverage Verified from tests/",
        ["Test File", "Focus"],
        [
            ["tests/test_sentence_buffer.py", "English and Arabic sentence flushing behavior."],
            ["tests/test_voice_normalizer.py", "Weather/search/unit normalization for spoken output."],
            ["tests/test_tts_prosody.py", "Punctuation cleanup, Arabic discourse punctuation, markdown stripping, and connector handling."],
            ["tests/test_llm_routing_guard.py", "Guarding Arabic/code-switched advice questions from being misrouted."],
            ["desktop/src/__tests__/*.test.ts", "Desktop protocol and store tests exist; not run in this Python pytest pass."],
        ],
    )
    para(doc, "4.16 Performance Analysis", style="Heading 2")
    para(doc, "The codebase includes timing and metrics hooks in core/metrics.py, orchestration stage timers, wake inference summaries, STT metadata, and command latency logging. This documentation pass did not find a formal benchmark dataset or reproducible performance report, so previous exact accuracy, false-acceptance, and latency claims should be treated as placeholders until measured under a documented protocol.")
    para(doc, "4.17 Graphs and Tables", style="Heading 2")
    para(doc, "No verified empirical graphs are included in this stage. Future formatting passes may add graphs only if backed by collected data, test logs, or benchmark scripts stored with the project.")
    doc.add_page_break()


def add_chapter_5(doc):
    para(doc, "Chapter 5: Conclusions & Future Work", style="Heading 1")
    for heading, body in [
        ("5.1 Final Conclusions", "JARVIS demonstrates that a bilingual Windows desktop assistant can be built as a local-first system by integrating wake-word detection, VAD, STT, NLU routing, local LLM fallback, TTS, and Windows automation modules."),
        ("5.2 How Objectives Were Achieved", "The project achieved its main objectives through unified wake-word configuration, bilingual STT validation, code-switch routing, modular Windows automation handlers, local-first runtime choices, policy/confirmation controls, and package-level separation of concerns."),
    ]:
        para(doc, heading, style="Heading 2")
        para(doc, body)
    para(doc, "5.3 Limitations", style="Heading 2")
    for item in [
        "No formal large-scale benchmark for STT accuracy, wake-word false acceptance, or end-to-end latency was found in the repository.",
        "Some integrations depend on optional packages, Windows-specific APIs, external services, or user configuration.",
        "Egyptian Arabic remains difficult because of dialect variation, spelling inconsistency, and code-switching.",
        "The implementation is Windows-specific and would require substantial adapter replacement for macOS or Linux.",
    ]:
        bullet(doc, item)
    para(doc, "5.4 Future Improvements", style="Heading 2")
    for item in [
        "Create a reproducible evaluation dataset for wake-word detection, VAD, STT, command routing, latency, and bilingual command success rate.",
        "Expand Egyptian Arabic command vocabulary and collect representative utterances from different speakers.",
        "Document and test the desktop UI as a user-facing configuration and monitoring surface.",
        "Add benchmark scripts that export tables and charts directly into the documentation appendices.",
        "Strengthen privacy documentation by separating local-only paths from optional network-backed paths.",
    ]:
        bullet(doc, item)
    doc.add_page_break()


def add_references_and_appendices(doc):
    para(doc, "References", style="Heading 1")
    for ref in [
        "[R1] Radford et al., Robust Speech Recognition via Large-Scale Weak Supervision, ICML/PMLR, 2023. https://proceedings.mlr.press/v202/radford23a.html",
        "[R2] OpenWakeWord project repository. https://github.com/dscripka/openWakeWord",
        "[R3] Silero VAD project repository. https://github.com/snakers4/silero-vad",
        "[R4] Faster-Whisper project repository. https://github.com/SYSTRAN/faster-whisper",
        "[R5] Ollama project documentation. https://ollama.com",
        "[R6] Open-Meteo weather API documentation. https://open-meteo.com",
        "[R7] Repository source files: main.py, core/, audio/, nlp/, llm/, os_control/, tools/, tests/.",
    ]:
        para(doc, ref)
    doc.add_page_break()

    para(doc, "Appendix A - Source Code Listings", style="Heading 1")
    para(doc, "Long source-code listings are not duplicated in the main chapters. The following table identifies the primary files to consult for implementation review.")
    add_table(
        doc,
        "Table A.1 - Source-Code Listing Reference",
        ["Area", "Files"],
        [
            ["Entry point", "main.py"],
            ["Orchestration", "core/orchestrator.py"],
            ["Configuration", "core/config.py"],
            ["Routing", "core/command_router.py, core/command_parser.py"],
            ["Audio", "audio/wake_word.py, audio/vad.py, audio/stt.py, audio/tts.py"],
            ["Memory", "core/session_memory.py, core/memory_store.py"],
            ["Windows automation", "os_control/*.py"],
            ["Tests", "tests/*.py"],
        ],
    )
    para(doc, "Appendix B - Configuration Reference", style="Heading 1")
    para(doc, "The full configuration surface is defined in core/config.py and .env.example. Sensitive values such as API keys belong in .env and should not be submitted publicly.")
    para(doc, "Appendix C - Installation Guide", style="Heading 1")
    add_table(
        doc,
        "Table B.1 - Installation Guide",
        ["Step", "Command / Action"],
        [
            ["1", "python -m pip install -r requirements.txt"],
            ["2", "copy .env.example .env"],
            ["3", "Edit .env and add optional keys only when needed."],
            ["4", "Start Ollama if using local LLM fallback: ollama serve"],
            ["5", "Run the assistant: python main.py"],
            ["6", "Optional health check: python core/doctor.py"],
        ],
    )
    para(doc, "Appendix D - User Manual", style="Heading 1")
    para(doc, "Start the runtime, wait for the ready state, say the wake word, speak a supported command or question, and listen for the response. Sensitive commands may require spoken PIN confirmation depending on configuration.")
    para(doc, "Appendix E - Testing Scenarios", style="Heading 1")
    for item in [
        "Run python -m pytest -q for automated Python tests.",
        "Test wake word with English and Egyptian Arabic pronunciations.",
        "Test no-speech behavior after wake activation.",
        "Test safe commands such as opening an app or reading system information.",
        "Test confirmation flow for sensitive system actions without executing destructive operations.",
    ]:
        bullet(doc, item)
    para(doc, "Appendix F - Glossary", style="Heading 1")
    for item in [
        "Parser fast path: deterministic command parsing before semantic or LLM fallback.",
        "Code-switch routing: handling commands that mix Egyptian Arabic and English terms.",
        "Follow-up window: a dialogue state where the next utterance may be accepted without repeating the wake word.",
        "Policy profile: configuration that controls which classes of actions are allowed.",
    ]:
        bullet(doc, item)
    para(doc, "Appendix G - Contribution Table", style="Heading 1")
    add_table(
        doc,
        "Table F.1 - Contribution Table Placeholder",
        ["Team Member", "Contribution", "Evidence / Files"],
        [
            ["[Name]", "[Role to be completed]", "[Files or tasks]"],
            ["[Name]", "[Role to be completed]", "[Files or tasks]"],
            ["[Name]", "[Role to be completed]", "[Files or tasks]"],
        ],
    )


def build():
    acknowledgement = extract_acknowledgement()
    doc = Document()
    set_margins(doc.sections[0])
    add_page_number(doc.sections[0])
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(11)
    for name, size in [("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 12)]:
        doc.styles[name].font.name = "Times New Roman"
        doc.styles[name].font.size = Pt(size)

    add_preliminaries(doc, acknowledgement)
    add_chapter_1(doc)
    add_chapter_2(doc)
    add_chapter_3(doc)
    add_chapter_4(doc)
    add_chapter_5(doc)
    add_references_and_appendices(doc)

    for section in doc.sections:
        set_margins(section)
    doc.save(str(DOC_PATH))


if __name__ == "__main__":
    build()
    print(f"Updated {DOC_PATH}")
