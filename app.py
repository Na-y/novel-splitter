import io
import os
import zipfile
import regex as re
import streamlit as st
from docx import Document

# Pre-compile regex patterns ported directly from your PHP function
CLEAN_PATTERN = re.compile(
    r'^[\s\p{Zs}\p{Cf}\uFEFF]+|[\s\p{Zs}\p{Cf}\uFEFF]+$',
    flags=re.UNICODE
)

NUM_PATTERN = r'[0-9၀-၉]+'
SEP_PATTERN = r'[\s\p{Zs}\p{P}\u200B-\u200D\u104A\u104B\-–—\(\[\{:.]*'

# Track A: Burmese & English explicit prefixes (e.g. အခန်း ၁, အပိုင်း ၂, Chapter 1, Ch. 12)
TRACK_A_PATTERN = re.compile(
    rf'^[\p{{S}}\p{{P}}\p{{Cf}}\s]*(?:အခန်း|အပိုင်း|Chapter|Ch\.?){SEP_PATTERN}({NUM_PATTERN})',
    flags=re.UNICODE | re.IGNORECASE
)

# Track B: Chinese / East Asian prefixes (e.g. 第1章, 第 12 话, 第3节)
TRACK_B_PATTERN = re.compile(
    rf'^[\p{{S}}\p{{P}}\s]*第{SEP_PATTERN}({NUM_PATTERN}){SEP_PATTERN}(?:章|話|话|節|节)',
    flags=re.UNICODE
)


def detect_chapter_heading(line: str) -> tuple[bool, str]:
    """
    Port of PHP detect_chapter_heading($line, &$detected_title).
    Returns (is_chapter: bool, detected_title: str).
    """
    # Strip invisible characters, zero-width spaces, and leading/trailing whitespace
    clean = CLEAN_PATTERN.sub('', line)

    if not clean or len(clean) > 180:
        return False, ''

    # Track A: Burmese & English
    match_a = TRACK_A_PATTERN.match(clean)
    if match_a:
        return True, f"အခန်း {match_a.group(1)}"

    # Track B: Chinese / East Asian
    match_b = TRACK_B_PATTERN.match(clean)
    if match_b:
        num = match_b.group(1)
        return True, f"အခန်း {num} (第{num}章)"

    return False, ''


def extract_lines_from_file(uploaded_file) -> list[str]:
    """Reads lines from uploaded .txt or .docx file."""
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext == '.txt':
        raw_bytes = uploaded_file.read()
        # Try UTF-8 first (with BOM support), fallback to UTF-16 if needed
        try:
            content = raw_bytes.decode('utf-8-sig')
        except UnicodeDecodeError:
            content = raw_bytes.decode('utf-16')
        return content.splitlines()
    elif ext == '.docx':
        doc = Document(uploaded_file)
        return [p.text for p in doc.paragraphs]
    return []


def parse_chapters(lines: list[str], normalize_heading: bool = False) -> tuple[list[str], list[dict]]:
    """
    Splits lines into chapters using detect_chapter_heading.
    Returns (preamble_lines, chapters_list).
    Each item in chapters_list is {'title': detected_title, 'raw_heading': line, 'lines': [...]}.
    """
    preamble = []
    chapters = []
    current_chapter = None

    for line in lines:
        is_heading, detected_title = detect_chapter_heading(line)
        if is_heading:
            if current_chapter is not None:
                chapters.append(current_chapter)
            heading_line = detected_title if normalize_heading else line
            current_chapter = {
                'title': detected_title,
                'raw_heading': line.strip(),
                'lines': [heading_line]
            }
        else:
            if current_chapter is None:
                preamble.append(line)
            else:
                current_chapter['lines'].append(line)

    if current_chapter is not None:
        chapters.append(current_chapter)

    return preamble, chapters


def build_docx_bytes(lines: list[str]) -> bytes:
    """Creates a .docx file in memory from a list of paragraph strings."""
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# --- Streamlit UI ---
st.set_page_config(page_title="Novel Chapter Splitter", page_icon="📚", layout="wide")
st.title("📚 Novel Chapter Splitter (.txt & .docx)")

uploaded_file = st.file_uploader("Upload a novel file (.txt or .docx)", type=["txt", "docx"])

col1, col2, col3 = st.columns(3)
with col1:
    chapters_per_file = st.number_input("Chapters per split file", min_value=1, value=30, step=1)
with col2:
    output_format = st.selectbox("Output File Format", ["Match Input", ".txt", ".docx"])
with col3:
    normalize_heading = st.checkbox(
        "Replace original heading with normalized title",
        value=False,
        help="Replaces the original chapter line with the standardized $detected_title (e.g., 'အခန်း ၁')"
    )

include_preamble = st.checkbox(
    "Keep text before Chapter 1 (synopsis/title) at the top of File 1",
    value=True
)

if uploaded_file is not None:
    base_name, input_ext = os.path.splitext(uploaded_file.name)
    out_ext = input_ext.lower() if output_format == "Match Input" else output_format

    lines = extract_lines_from_file(uploaded_file)
    preamble, chapters = parse_chapters(lines, normalize_heading=normalize_heading)

    total_chapters = len(chapters)

    if total_chapters == 0:
        st.warning("No chapters matched the detection rules in this file.")
    else:
        # Chunk chapters into groups of `chapters_per_file`
        chunks = [
            chapters[i:i + chapters_per_file]
            for i in range(0, total_chapters, chapters_per_file)
        ]
        chunk_Counts = [len(c) for c in chunks]

        st.success(
            f"Detected **{total_chapters}** chapters. "
            f"Splitting at **{chapters_per_file}** chapters/file produces **{len(chunks)}** files: "
            f"`{', '.join(map(str, chunk_Counts))}` chapters."
        )

        # Build ZIP archive in memory
        zip_buffer = io.BytesIO()
        split_summary = []

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for idx, chunk in enumerate(chunks, start=1):
                start_ch = (idx - 1) * chapters_per_file + 1
                end_ch = start_ch + len(chunk) - 1
                file_name = f"{base_name}_part{idx:02d}_ch{start_ch}-{end_ch}{out_ext}"

                combined_lines = []
                if idx == 1 and include_preamble and any(p.strip() for p in preamble):
                    combined_lines.extend(preamble)
                    combined_lines.append("")

                for ch_idx, ch in enumerate(chunk):
                    if ch_idx > 0:
                        combined_lines.append("")  # Blank line between chapters
                    combined_lines.extend(ch['lines'])

                if out_ext == ".txt":
                    file_bytes = "\n".join(combined_lines).encode("utf-8")
                else:
                    file_bytes = build_docx_bytes(combined_lines)

                zip_file.writestr(file_name, file_bytes)

                split_summary.append({
                    "File #": idx,
                    "Filename": file_name,
                    "Chapter Range": f"{start_ch} – {end_ch}",
                    "Chapter Count": len(chunk),
                    "First Detected Heading": f"{chunk[0]['title']} ({chunk[0]['raw_heading'][:40]})",
                    "Last Detected Heading": f"{chunk[-1]['title']} ({chunk[-1]['raw_heading'][:40]})"
                })

        st.download_button(
            label=f"⬇️ Download All {len(chunks)} Split Files (.zip)",
            data=zip_buffer.getvalue(),
            file_name=f"{base_name}_split_{chapters_per_file}ch.zip",
            mime="application/zip",
            type="primary"
        )

        st.subheader("Split Files Preview")
        st.dataframe(split_summary, use_container_width=True, hide_index=True)

        with st.expander("View All Detected Chapter Headings"):
            heading_rows = [
                {
                    "Index": i,
                    "Detected Title": ch["title"],
                    "Original Line": ch["raw_heading"]
                }
                for i, ch in enumerate(chapters, start=1)
            ]
            st.dataframe(heading_rows, use_container_width=True, hide_index=True)
