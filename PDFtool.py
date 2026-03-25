import streamlit as st
import pymupdf as fitz
import io
from PIL import Image
import uuid

# --- Page Configuration ---
st.set_page_config(page_title="Ultimate PDF Editor", page_icon="📚", layout="wide")

# --- Initialize Session State ---
if "source_streams" not in st.session_state:
    st.session_state.source_streams = {}
if "pages_workbench" not in st.session_state:
    st.session_state.pages_workbench = []
if "final_pdf_bytes" not in st.session_state:
    st.session_state.final_pdf_bytes = None
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = str(uuid.uuid4())

# --- Callback: Fast Page Jumping ---
def move_page(old_idx, widget_key):
    new_idx = st.session_state[widget_key] - 1
    if new_idx != old_idx:
        item = st.session_state.pages_workbench.pop(old_idx)
        st.session_state.pages_workbench.insert(new_idx, item)
        st.session_state.final_pdf_bytes = None

# --- NEW Callback: File-Level Batch Operations ---
def move_file_group(file_id, direction):
    """Moves all pages of a specific file up or down as a contiguous block."""
    # Step 1: Find the current relative order of unique files in the workbench
    current_order = []
    for p in st.session_state.pages_workbench:
        if p["source_file_id"] not in current_order:
            current_order.append(p["source_file_id"])

    # Step 2: Determine new position
    idx = current_order.index(file_id)
    new_idx = idx + direction

    if 0 <= new_idx < len(current_order):
        # Swap the order of files
        current_order[idx], current_order[new_idx] = current_order[new_idx], current_order[idx]

        # Step 3: Rebuild the workbench exactly matching the new file order
        # This automatically gathers and regroups scattered pages from the same file
        new_workbench = []
        for fid in current_order:
            new_workbench.extend([p for p in st.session_state.pages_workbench if p["source_file_id"] == fid])

        st.session_state.pages_workbench = new_workbench
        st.session_state.final_pdf_bytes = None

# --- Core Logic Functions ---
def load_new_files_to_workbench(uploaded_files):
    if not uploaded_files:
        return

    with st.spinner("New file detected, parsing automatically..."):
        new_pages_added_count = 0
        for uploaded_file in uploaded_files:
            file_id = uploaded_file.name
            
            if file_id in st.session_state.source_streams:
                continue

            pdf_bytes = uploaded_file.read()
            st.session_state.source_streams[file_id] = pdf_bytes
            
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(dpi=72)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                st.session_state.pages_workbench.append({
                    "unique_id": str(uuid.uuid4()),
                    "source_file_id": file_id,
                    "original_index": page_num,
                    "image": img,
                    "rotation": 0
                })
                new_pages_added_count += 1
            doc.close()
        
        if new_pages_added_count > 0:
            st.toast(f"✅ Parsed successfully! Added {new_pages_added_count} pages.")
            st.session_state.final_pdf_bytes = None

def clear_all():
    st.session_state.source_streams = {}
    st.session_state.pages_workbench = []
    st.session_state.final_pdf_bytes = None
    st.session_state.uploader_key = str(uuid.uuid4())
    st.rerun()


# ==========================================
# Sidebar: Control Panel
# ==========================================
with st.sidebar:
    st.header("⚙️ Control Panel")
    st.markdown("Upload files and adjust settings here.")
    st.divider()
    
    st.subheader("👁️ View Settings")
    cols_per_row = st.slider("Cards per row", min_value=2, max_value=8, value=5, step=1)
    
    st.divider()
    
    st.subheader("📥 1. Upload Files")
    uploaded_files = st.file_uploader(
        "Select or drag multiple PDFs", 
        type="pdf", 
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=st.session_state.uploader_key
    )

    if uploaded_files:
        new_files_to_process = [f for f in uploaded_files if f.name not in st.session_state.source_streams]
        if new_files_to_process:
            load_new_files_to_workbench(new_files_to_process)
            st.rerun()
            
    if st.button("🗑️ Clear Workbench", use_container_width=True):
        clear_all()

    st.divider()

    st.subheader("💾 2. Export")
    export_filename = st.text_input("Export Filename", value="Merged_Edited.pdf")
    
    is_workbench_empty = len(st.session_state.pages_workbench) == 0
    
    if st.button(
        "🚀 Generate Final PDF", 
        type="primary", 
        use_container_width=True, 
        disabled=is_workbench_empty,
        help="Please upload files above first" if is_workbench_empty else "Click to generate the merged PDF"
    ):
        with st.spinner("Merging..."):
            out_doc = fitz.open()
            opened_docs = {}
            try:
                for page_info in st.session_state.pages_workbench:
                    file_id = page_info["source_file_id"]
                    orig_idx = page_info["original_index"]
                    relative_rotation = page_info["rotation"]
                    
                    if file_id not in opened_docs:
                        source_stream = st.session_state.source_streams[file_id]
                        opened_docs[file_id] = fitz.open(stream=source_stream, filetype="pdf")
                            
                    src_doc = opened_docs[file_id]
                    out_doc.insert_pdf(src_doc, from_page=orig_idx, to_page=orig_idx)
                    
                    new_page = out_doc[-1]
                    if relative_rotation != 0:
                        current_rot = new_page.rotation
                        new_page.set_rotation((current_rot + relative_rotation) % 360)
                
                st.session_state.final_pdf_bytes = out_doc.write()
            except Exception as e:
                st.error(f"Error occurred: {e}")
            finally:
                for doc in opened_docs.values():
                    doc.close()
                out_doc.close()
    
    if st.session_state.final_pdf_bytes:
        st.success("✅ Generated successfully!")
        st.download_button(
            label="📥 Download PDF",
            data=st.session_state.final_pdf_bytes,
            file_name=export_filename,
            mime="application/pdf",
            use_container_width=True
        )


# ==========================================
# Main Interface: Visual Workbench
# ==========================================
st.title("📚 Ultimate PDF Editor")
st.markdown("👈 **Drag and drop files on the left.** Sort, delete, and rotate pages below.")
st.divider()

st.subheader(f"🛠️ Workbench ({len(st.session_state.pages_workbench)} pages)")

if not st.session_state.pages_workbench:
    st.info("Workbench is empty. Please upload PDF files in the left panel.")
else:
    # --- NEW: File-Level Batch Operations Menu ---
    with st.expander("📦 Batch Operations: Manage by File", expanded=False):
        unique_files = []
        for p in st.session_state.pages_workbench:
            if p["source_file_id"] not in unique_files:
                unique_files.append(p["source_file_id"])

        for idx, f_id in enumerate(unique_files):
            file_pages_count = sum(1 for p in st.session_state.pages_workbench if p["source_file_id"] == f_id)
            
            col_name, col_up, col_down, col_rm = st.columns([5, 2, 2, 2])
            
            col_name.markdown(f"📄 **{f_id}** <span style='color:gray; font-size: 0.9em;'>({file_pages_count} pages)</span>", unsafe_allow_html=True)
            
            if col_up.button("⬆️ Move Up", key=f"up_{f_id}", disabled=(idx == 0), use_container_width=True):
                move_file_group(f_id, -1)
                st.rerun()
                
            if col_down.button("⬇️ Move Down", key=f"down_{f_id}", disabled=(idx == len(unique_files) - 1), use_container_width=True):
                move_file_group(f_id, 1)
                st.rerun()
                
            # Allow bulk deletion of an entire file's pages
            if col_rm.button("❌ Remove", key=f"rm_{f_id}", use_container_width=True):
                st.session_state.pages_workbench = [p for p in st.session_state.pages_workbench if p["source_file_id"] != f_id]
                st.session_state.final_pdf_bytes = None
                st.rerun()
                
    st.divider()

    # --- Page Cards Rendering ---
    for i in range(0, len(st.session_state.pages_workbench), cols_per_row):
        cols = st.columns(cols_per_row)
        for j in range(cols_per_row):
            idx = i + j
            if idx < len(st.session_state.pages_workbench):
                page_info = st.session_state.pages_workbench[idx]
                current_key = page_info["unique_id"]
                col = cols[j]
                
                with col:
                    with st.container(border=True):
                        
                        # 1. Top Navigation & Metadata
                        top_cols = st.columns([1.5, 2.5, 1.2])
                        
                        with top_cols[0]:
                            st.markdown(f"**Page {idx + 1}**")
                            
                        with top_cols[1]:
                            st.markdown(
                                f"<div style='white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: 0.8em; color: gray; padding-top: 0.2rem;' title='{page_info['source_file_id']} (Orig P{page_info['original_index'] + 1})'>"
                                f"{page_info['source_file_id']}"
                                f"</div>", 
                                unsafe_allow_html=True
                            )
                            
                        with top_cols[2]:
                            with st.popover("🎯", help="Jump to page"):
                                st.number_input(
                                    "Jump to page number:", 
                                    min_value=1, 
                                    max_value=len(st.session_state.pages_workbench), 
                                    value=idx + 1, 
                                    key=f"jump_{current_key}",
                                    on_change=move_page,
                                    args=(idx, f"jump_{current_key}")
                                )
                        
                        # 2. Middle: Page Preview
                        display_img = page_info["image"].rotate(-page_info["rotation"], expand=True)
                        st.image(display_img, use_container_width=True)
                        
                        # 3. Bottom: Action Buttons
                        btn_cols = st.columns(4, gap="small")
                        
                        if btn_cols[0].button("⬅️", key=f"left_{current_key}", disabled=(idx == 0), help="Move left"):
                            st.session_state.pages_workbench.insert(idx - 1, st.session_state.pages_workbench.pop(idx))
                            st.session_state.final_pdf_bytes = None
                            st.rerun()
                            
                        if btn_cols[1].button("➡️", key=f"right_{current_key}", disabled=(idx == len(st.session_state.pages_workbench) - 1), help="Move right"):
                            st.session_state.pages_workbench.insert(idx + 1, st.session_state.pages_workbench.pop(idx))
                            st.session_state.final_pdf_bytes = None
                            st.rerun()
                            
                        if btn_cols[2].button("🔄", key=f"rot_{current_key}", help="Rotate 90° clockwise"):
                            st.session_state.pages_workbench[idx]["rotation"] = (st.session_state.pages_workbench[idx]["rotation"] + 90) % 360
                            st.session_state.final_pdf_bytes = None
                            st.rerun()
                            
                        if btn_cols[3].button("❌", key=f"del_{current_key}", help="Delete this page"):
                            st.session_state.pages_workbench.pop(idx)
                            st.session_state.final_pdf_bytes = None
                            st.rerun()
