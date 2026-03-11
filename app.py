import io
import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

import streamlit as st

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt
except ImportError:  # pragma: no cover - handled in UI
    Presentation = None
    RGBColor = None
    Inches = None
    Pt = None


APP_NAME = "ai_presentation_generator"
DEFAULT_MODEL = "gemini-2.5-flash"
API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
UNSPLASH_RANDOM_URL = "https://api.unsplash.com/photos/random"
OUTPUT_DIR = Path("generated_presentations")
PRESENTATION_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "title": {"type": "STRING"},
        "subtitle": {"type": "STRING"},
        "theme": {
            "type": "OBJECT",
            "properties": {
                "primary": {"type": "STRING"},
                "secondary": {"type": "STRING"},
                "accent": {"type": "STRING"},
            },
            "required": ["primary", "secondary", "accent"],
        },
        "slides": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "title": {"type": "STRING"},
                    "bullets": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"},
                    },
                    "speaker_notes": {"type": "STRING"},
                },
                "required": ["title", "bullets", "speaker_notes"],
            },
        },
        "closing_message": {"type": "STRING"},
    },
    "required": ["title", "subtitle", "theme", "slides", "closing_message"],
}


st.set_page_config(
    page_title="AI Presentation Generator",
    page_icon=":material/slideshow:",
    layout="wide",
)


@st.cache_data(show_spinner=False)
def load_dotenv(path: str = ".env") -> dict[str, str]:
    # Load simple KEY=value pairs without depending on another package.
    env_path = Path(path)
    loaded: dict[str, str] = {}
    if not env_path.exists():
        return loaded

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)
        loaded[key] = value

    return loaded


def get_api_key() -> str | None:
    load_dotenv()
    return os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")


def get_unsplash_key() -> str | None:
    load_dotenv()
    return os.getenv("UNSPLASH_ACCESS_KEY")


def build_prompt(topic: str, audience: str, tone: str, slide_count: int) -> str:
    # Keep the prompt explicit so Gemini reliably returns presentation-ready JSON.
    return f"""
Create a professional presentation outline as JSON.

Topic: {topic}
Audience: {audience}
Tone: {tone}
Requested content slides: {slide_count}

Instructions:
- Return valid JSON only.
- Create a strong presentation title and a one-sentence subtitle.
- Provide exactly {slide_count} content slides in the slides array.
- Each slide needs a concise title, 3 to 4 bullets, and speaker_notes.
- Bullets should be short, specific, and presentation-ready.
- Use a practical, informative tone for a live presentation.
- Theme colors must be hex values like #1F3A5F.
- closing_message should be a short ending line for the final slide.
""".strip()


@st.cache_data(show_spinner=False)
def call_gemini(prompt: str, api_key: str, model: str) -> dict[str, Any]:
    # Ask Gemini for structured JSON so the app can safely turn the result into slides.
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.7,
            "responseMimeType": "application/json",
            "responseSchema": PRESENTATION_SCHEMA,
        },
    }

    request = urllib.request.Request(
        API_URL.format(model=model),
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "X-goog-api-key": api_key,
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Gemini API request failed ({exc.code}). {details}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Network error while contacting Gemini: {exc.reason}") from exc

    response_json = json.loads(raw)
    parts = response_json.get("candidates", [{}])[0].get("content", {}).get("parts", [])
    text = "".join(part.get("text", "") for part in parts).strip()
    if not text:
        raise RuntimeError("Gemini returned an empty response.")

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, re.S)
        if not match:
            raise RuntimeError("Gemini did not return valid JSON.")
        return json.loads(match.group(0))


def sanitize_hex(value: Any, fallback: str) -> str:
    if isinstance(value, str) and re.fullmatch(r"#[0-9A-Fa-f]{6}", value.strip()):
        return value.strip().upper()
    return fallback


def normalize_theme(raw_theme: dict[str, Any] | None) -> dict[str, str]:
    theme = raw_theme or {}
    return {
        "primary": sanitize_hex(theme.get("primary"), "#1F3A5F"),
        "secondary": sanitize_hex(theme.get("secondary"), "#F5F1EA"),
        "accent": sanitize_hex(theme.get("accent"), "#CB6D43"),
    }


def with_referral(url: str | None) -> str:
    if not url:
        return ""
    separator = "&" if "?" in url else "?"
    return f"{url}{separator}utm_source={APP_NAME}&utm_medium=referral"


def fetch_json(url: str, headers: dict[str, str]) -> dict[str, Any]:
    request = urllib.request.Request(url, headers=headers, method="GET")
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_unsplash_photo(query: str, access_key: str) -> dict[str, str] | None:
    # Unsplash requires using returned image URLs directly and preserving attribution links.
    params = urllib.parse.urlencode(
        {
            "query": query,
            "orientation": "landscape",
            "content_filter": "high",
        }
    )
    headers = {"Authorization": f"Client-ID {access_key}", "Accept-Version": "v1"}

    try:
        photo = fetch_json(f"{UNSPLASH_RANDOM_URL}?{params}", headers)
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Unsplash request failed ({exc.code}). {details}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Network error while contacting Unsplash: {exc.reason}") from exc

    urls = photo.get("urls", {})
    user = photo.get("user", {})
    links = photo.get("links", {})
    profile = user.get("links", {}).get("html")

    return {
        "image_url": urls.get("regular", ""),
        "download_location": links.get("download_location", ""),
        "photographer": user.get("name", "Unsplash photographer"),
        "profile_url": with_referral(profile),
        "photo_url": with_referral(links.get("html")),
    }


def trigger_unsplash_download(download_location: str, access_key: str) -> None:
    if not download_location:
        return

    headers = {"Authorization": f"Client-ID {access_key}", "Accept-Version": "v1"}
    try:
        fetch_json(download_location, headers)
    except Exception:
        # Export should still succeed if usage tracking fails.
        return


def fetch_image_bytes(url: str) -> bytes:
    if not url:
        return b""

    request = urllib.request.Request(url, headers={"User-Agent": APP_NAME}, method="GET")
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read()


def normalize_deck(data: dict[str, Any], topic: str, slide_count: int) -> dict[str, Any]:
    # Clean the model output and fill any gaps so export always has a complete deck.
    slides: list[dict[str, Any]] = []
    for item in data.get("slides", []):
        bullets = [str(bullet).strip() for bullet in item.get("bullets", []) if str(bullet).strip()]
        if not bullets:
            continue
        slides.append(
            {
                "title": str(item.get("title") or "Untitled Slide").strip(),
                "bullets": bullets[:4],
                "speaker_notes": str(item.get("speaker_notes") or "").strip(),
                "image": None,
            }
        )

    while len(slides) < slide_count:
        slides.append(
            {
                "title": f"Key Point {len(slides) + 1}",
                "bullets": [
                    f"Expand on an important angle of {topic}",
                    "Add a concrete example or data point",
                    "Close with a takeaway for the audience",
                ],
                "speaker_notes": "Use this slide as a fallback if the model returns fewer slides than requested.",
                "image": None,
            }
        )

    return {
        "title": str(data.get("title") or topic.title()).strip(),
        "subtitle": str(data.get("subtitle") or f"An AI-generated presentation about {topic}").strip(),
        "theme": normalize_theme(data.get("theme")),
        "slides": slides[:slide_count],
        "closing_message": str(data.get("closing_message") or "Questions and discussion").strip(),
    }


def enrich_deck_with_unsplash(deck: dict[str, Any], topic: str, access_key: str | None) -> dict[str, Any]:
    if not access_key:
        return deck

    for slide in deck["slides"]:
        query = f"{topic} {slide['title']}"
        try:
            slide["image"] = fetch_unsplash_photo(query, access_key)
        except Exception:
            slide["image"] = None
    return deck


def hex_to_rgb(value: str) -> RGBColor:
    cleaned = value.lstrip("#")
    return RGBColor(int(cleaned[0:2], 16), int(cleaned[2:4], 16), int(cleaned[4:6], 16))


def apply_background(slide, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(color)


def add_footer_text(slide, text: str, color: str) -> None:
    textbox = slide.shapes.add_textbox(Inches(0.6), Inches(7.0), Inches(12.0), Inches(0.25))
    paragraph = textbox.text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.size = Pt(8)
    paragraph.font.color.rgb = hex_to_rgb(color)


def build_presentation(deck: dict[str, Any], unsplash_key: str | None) -> bytes:
    # Build a real PowerPoint file entirely in memory before offering download/save.
    if Presentation is None:
        raise RuntimeError("python-pptx is not installed. Install dependencies from requirements.txt first.")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    theme = deck["theme"]
    primary = theme["primary"]
    secondary = theme["secondary"]
    accent = theme["accent"]

    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    apply_background(title_slide, primary)
    title_box = title_slide.shapes.title
    subtitle_box = title_slide.placeholders[1]
    title_box.text = deck["title"]
    subtitle_box.text = deck["subtitle"]
    title_paragraph = title_box.text_frame.paragraphs[0]
    title_paragraph.font.size = Pt(28)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = hex_to_rgb("#FFFFFF")
    subtitle_paragraph = subtitle_box.text_frame.paragraphs[0]
    subtitle_paragraph.font.size = Pt(16)
    subtitle_paragraph.font.color.rgb = hex_to_rgb("#F6F0E8")

    for slide_data in deck["slides"]:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        apply_background(slide, secondary)

        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.45), Inches(12.0), Inches(0.65))
        title_paragraph = title_box.text_frame.paragraphs[0]
        title_paragraph.text = slide_data["title"]
        title_paragraph.font.size = Pt(24)
        title_paragraph.font.bold = True
        title_paragraph.font.color.rgb = hex_to_rgb(primary)

        body_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.35), Inches(5.4), Inches(4.7))
        body_frame = body_box.text_frame
        body_frame.word_wrap = True
        for index, bullet in enumerate(slide_data["bullets"]):
            paragraph = body_frame.paragraphs[0] if index == 0 else body_frame.add_paragraph()
            paragraph.text = bullet
            paragraph.level = 0
            paragraph.font.size = Pt(20)
            paragraph.font.color.rgb = hex_to_rgb("#2E2A25")
            if index == 0:
                paragraph.space_after = Pt(10)

        image = slide_data.get("image")
        if image and image.get("image_url"):
            image_bytes = fetch_image_bytes(image["image_url"])
            if image_bytes:
                slide.shapes.add_picture(io.BytesIO(image_bytes), Inches(6.6), Inches(1.35), width=Inches(5.9), height=Inches(3.8))
                attribution = f"Photo: {image['photographer']} on Unsplash"
                add_footer_text(slide, attribution, primary)
                if unsplash_key:
                    trigger_unsplash_download(image.get("download_location", ""), unsplash_key)

        notes_text_frame = slide.notes_slide.notes_text_frame
        notes_text_frame.text = slide_data["speaker_notes"]
        if image and image.get("photographer"):
            notes_text_frame.text += (
                f"\n\nImage credit: {image['photographer']} | "
                f"Profile: {image.get('profile_url', '')} | Photo: {image.get('photo_url', '')}"
            )

        accent_shape = slide.shapes.add_shape(1, Inches(0.6), Inches(6.65), Inches(11.9), Inches(0.15))
        accent_shape.fill.solid()
        accent_shape.fill.fore_color.rgb = hex_to_rgb(accent)
        accent_shape.line.fill.background()

    closing_slide = prs.slides.add_slide(prs.slide_layouts[5])
    apply_background(closing_slide, primary)
    closing_title = closing_slide.shapes.title
    if closing_title is None:
        closing_title = closing_slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.5), Inches(1.2))
        title_frame = closing_title.text_frame
    else:
        title_frame = closing_title.text_frame
    title_frame.text = "Thank You"
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.font.size = Pt(28)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = hex_to_rgb("#FFFFFF")

    message_box = closing_slide.shapes.add_textbox(Inches(0.9), Inches(2.4), Inches(11.2), Inches(1.4))
    message_frame = message_box.text_frame
    message_frame.text = deck["closing_message"]
    message_paragraph = message_frame.paragraphs[0]
    message_paragraph.font.size = Pt(20)
    message_paragraph.font.color.rgb = hex_to_rgb("#F6F0E8")

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


def make_file_stem(topic: str) -> str:
    safe_topic = re.sub(r"[^A-Za-z0-9_-]+", "-", topic.strip().lower()).strip("-")
    return safe_topic or "presentation"


def save_presentation_file(pptx_bytes: bytes, topic: str) -> Path:
    # Save a copy on disk so the user gets an actual file in the project folder.
    OUTPUT_DIR.mkdir(exist_ok=True)
    file_path = OUTPUT_DIR / f"{make_file_stem(topic)}.pptx"
    file_path.write_bytes(pptx_bytes)
    return file_path


api_key = get_api_key()
unsplash_key = get_unsplash_key()

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=Source+Sans+3:wght@400;600&display=swap');

        :root {
            --bg: #f4f4ef;
            --ink: #1d1c1a;
            --muted: #56524b;
            --accent: #cb6d43;
            --accent-strong: #9c4e2d;
            --line: rgba(29, 28, 26, 0.12);
            --good-bg: #e3f1e8;
            --good-ink: #1f5a36;
            --warn-bg: #f6e7d7;
            --warn-ink: #8a4d17;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(203, 109, 67, 0.18), transparent 30%),
                radial-gradient(circle at bottom right, rgba(65, 94, 84, 0.14), transparent 26%),
                linear-gradient(135deg, #f8f7f2 0%, var(--bg) 48%, #ece7df 100%);
            color: var(--ink);
        }

        .block-container {
            max-width: 1180px;
            padding-top: 2.25rem;
            padding-bottom: 2rem;
        }

        h1, h2, h3 {
            font-family: "Space Grotesk", sans-serif;
            color: var(--ink);
        }

        p, li, div[data-testid="stMarkdownContainer"], label {
            font-family: "Source Sans 3", sans-serif;
        }

        .hero-shell, .panel-shell {
            background: linear-gradient(180deg, rgba(255,255,255,0.82), rgba(255,255,255,0.58));
            border: 1px solid var(--line);
            box-shadow: 0 24px 80px rgba(73, 58, 43, 0.10);
            border-radius: 28px;
        }

        .hero-shell {
            padding: 3.4rem 3rem;
            margin-bottom: 1.25rem;
        }

        .panel-shell {
            padding: 1.35rem 1.4rem;
            margin-top: 1rem;
        }

        .eyebrow {
            display: inline-block;
            padding: 0.45rem 0.8rem;
            border-radius: 999px;
            background: rgba(239, 213, 199, 0.62);
            color: #7e4327;
            font-size: 0.9rem;
            font-weight: 700;
            text-transform: uppercase;
        }

        .hero-title {
            font-size: clamp(2.8rem, 5vw, 5.4rem);
            line-height: 0.94;
            margin: 1rem 0;
            max-width: 9ch;
        }

        .hero-copy, .sub-copy {
            color: var(--muted);
            line-height: 1.55;
        }

        .chip-row {
            display: flex;
            flex-wrap: wrap;
            gap: 0.7rem;
            margin-top: 1.25rem;
        }

        .chip, .status-pill {
            display: inline-block;
            padding: 0.65rem 0.95rem;
            border-radius: 999px;
            font-weight: 600;
        }

        .chip {
            background: rgba(29, 28, 26, 0.05);
            border: 1px solid rgba(29, 28, 26, 0.08);
        }

        .status-pill {
            margin-top: 1rem;
            font-weight: 700;
            font-size: 0.95rem;
        }

        .status-good {
            background: var(--good-bg);
            color: var(--good-ink);
        }

        .status-warn {
            background: var(--warn-bg);
            color: var(--warn-ink);
        }

        .slide-card {
            background: rgba(255, 255, 255, 0.72);
            border: 1px solid rgba(29, 28, 26, 0.08);
            border-radius: 22px;
            padding: 1.1rem 1.2rem;
            margin-bottom: 0.85rem;
        }

        .credit-text {
            color: var(--muted);
            font-size: 0.92rem;
            margin-top: 0.6rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

status_class = "status-good" if api_key else "status-warn"
status_text = "Gemini API key loaded from .env" if api_key else "Add GEMINI_API_KEY or GOOGLE_API_KEY to .env"
image_status_class = "status-good" if unsplash_key else "status-warn"
image_status_text = "Unsplash images enabled" if unsplash_key else "Add UNSPLASH_ACCESS_KEY to include Unsplash images"

st.markdown(
    f"""
    <section class="hero-shell">
        <span class="eyebrow">AI Presentation Generator</span>
        <h1 class="hero-title">Generate your deck.</h1>
        <div class="chip-row">
            <span class="chip">Gemini</span>
            <span class="chip">Unsplash</span>
            <span class="chip">PowerPoint</span>
        </div>
        <div class="status-pill {status_class}">{status_text}</div>
        <div class="status-pill {image_status_class}">{image_status_text}</div>
    </section>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Presentation Setup")
    topic = st.text_input("Topic", placeholder="e.g. AI in healthcare")
    audience = st.text_input("Audience", value="Business stakeholders")
    tone = st.selectbox("Tone", ["Professional", "Educational", "Persuasive", "Executive summary"], index=0)
    slide_count = st.slider("Content slides", min_value=4, max_value=10, value=6)
    model = st.selectbox("Gemini model", [DEFAULT_MODEL, "gemini-2.0-flash", "gemini-1.5-flash"], index=0)
    use_unsplash = st.toggle("Include Unsplash images", value=bool(unsplash_key))
    generate = st.button("Generate presentation", use_container_width=True)

if "deck" not in st.session_state:
    st.session_state.deck = None
if "pptx_bytes" not in st.session_state:
    st.session_state.pptx_bytes = None
if "pptx_path" not in st.session_state:
    st.session_state.pptx_path = None
if "last_error" not in st.session_state:
    st.session_state.last_error = None

if generate:
    st.session_state.last_error = None
    st.session_state.deck = None
    st.session_state.pptx_bytes = None
    st.session_state.pptx_path = None

    if not topic.strip():
        st.session_state.last_error = "Enter a presentation topic first."
    elif not api_key:
        st.session_state.last_error = "No API key found. Add GEMINI_API_KEY or GOOGLE_API_KEY to your .env file."
    elif use_unsplash and not unsplash_key:
        st.session_state.last_error = "Unsplash images are enabled, but UNSPLASH_ACCESS_KEY is missing from .env."
    else:
        prompt = build_prompt(topic.strip(), audience.strip() or "General audience", tone, slide_count)
        with st.spinner("Generating slide outline with Gemini..."):
            try:
                raw_deck = call_gemini(prompt, api_key, model)
                deck = normalize_deck(raw_deck, topic.strip(), slide_count)
                deck = enrich_deck_with_unsplash(deck, topic.strip(), unsplash_key if use_unsplash else None)
                pptx_bytes = build_presentation(deck, unsplash_key if use_unsplash else None)
                pptx_path = save_presentation_file(pptx_bytes, topic.strip())
            except Exception as exc:  # pragma: no cover - UI path
                st.session_state.last_error = str(exc)
            else:
                st.session_state.deck = deck
                st.session_state.pptx_bytes = pptx_bytes
                st.session_state.pptx_path = str(pptx_path)

if st.session_state.last_error:
    st.error(st.session_state.last_error)

left_col, right_col = st.columns([1.2, 0.8], gap="large")

with left_col:
    deck = st.session_state.deck
    if deck:
        tabs = st.tabs(["Slide Preview", "JSON Output"])
        with tabs[0]:
            st.subheader(deck["title"])
            st.caption(deck["subtitle"])
            for index, slide in enumerate(deck["slides"], start=1):
                bullets_html = "".join(f"<li>{bullet}</li>" for bullet in slide["bullets"])
                st.markdown(
                    f"""
                    <div class="slide-card">
                        <strong>Slide {index}: {slide['title']}</strong>
                        <ul>{bullets_html}</ul>
                        <p class="sub-copy"><strong>Speaker notes:</strong> {slide['speaker_notes']}</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                image = slide.get("image")
                if image and image.get("image_url"):
                    st.image(image["image_url"], use_container_width=True)
                    photographer = image.get("photographer", "Unsplash photographer")
                    photo_url = image.get("photo_url", "")
                    profile_url = image.get("profile_url", "")
                    st.markdown(
                        f"<p class='credit-text'>Photo by <a href='{profile_url}' target='_blank'>{photographer}</a> on <a href='{photo_url}' target='_blank'>Unsplash</a></p>",
                        unsafe_allow_html=True,
                    )
            st.markdown(
                f"""
                <div class="slide-card">
                    <strong>Final slide</strong>
                    <p>{deck['closing_message']}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with tabs[1]:
            st.json(deck)
    else:
        st.info("Generate a presentation to preview the outline and create the PowerPoint file.")

with right_col:
    if st.session_state.pptx_bytes:
        safe_topic = make_file_stem(topic)
        st.download_button(
            "Download PowerPoint",
            data=st.session_state.pptx_bytes,
            file_name=f"{safe_topic}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
        st.success("PowerPoint file is ready.")
        if st.session_state.pptx_path:
                st.code(st.session_state.pptx_path, language="text")
    else:
        st.warning("No PowerPoint file yet.")

    if Presentation is None:
        st.error("python-pptx is not installed in this environment yet. Install requirements before exporting decks.")

st.caption(
    "Uses Gemini for structured slide JSON and Unsplash for optional slide images, with attribution shown in the app and exported notes."
)
